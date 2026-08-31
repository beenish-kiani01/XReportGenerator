```python
import os
import re
import traceback
import requests

from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv

from playwright.sync_api import sync_playwright

from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT


# ============================================================
# CONFIGURATION
# ============================================================

load_dotenv()

TOKEN = os.getenv("X_BEARER_TOKEN")

app = Flask(__name__)

SCREENSHOT_DIR = "screenshots"
REPORT_DIR = "reports"

os.makedirs(SCREENSHOT_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)


# ============================================================
# SEARCH X POSTS
# ============================================================

def search_posts(query, number_of_posts):

    if not TOKEN:
        raise Exception("X_BEARER_TOKEN is missing.")

    url = "https://api.x.com/2/tweets/search/recent"

    headers = {
        "Authorization": f"Bearer {TOKEN}"
    }

    params = {
        "query": f"{query} -is:retweet",
        "max_results": max(
            10,
            min(int(number_of_posts), 100)
        ),
        "tweet.fields": "created_at,author_id",
        "expansions": "author_id",
        "user.fields": "username"
    }

    response = requests.get(
        url,
        headers=headers,
        params=params,
        timeout=30
    )

    print("X API STATUS:", response.status_code)

    if response.status_code != 200:
        print("X API RESPONSE:", response.text)
        raise Exception(
            f"X API request failed: {response.status_code}"
        )

    result = response.json()

    users = {
        user["id"]: user["username"]
        for user in result.get(
            "includes",
            {}
        ).get(
            "users",
            []
        )
    }

    posts = []

    for tweet in result.get("data", []):

        username = users.get(
            tweet.get("author_id"),
            "Unknown"
        )

        posts.append({
            "url": (
                f"https://x.com/"
                f"{username}/status/"
                f"{tweet['id']}"
            ),
            "handler_id": f"@{username}",
            "created_at": tweet.get(
                "created_at",
                ""
            )
        })

    posts.sort(
        key=lambda x: x.get(
            "created_at",
            ""
        ),
        reverse=True
    )

    return posts[:int(number_of_posts)]


# ============================================================
# DIRECT X POST URL
# ============================================================

def get_post_from_url(url):

    pattern = (
        r"(?:https?://)?"
        r"(?:www\.)?"
        r"(?:x\.com|twitter\.com)/"
        r"([^/]+)/status/(\d+)"
    )

    match = re.search(
        pattern,
        url
    )

    if not match:
        raise Exception(
            "Invalid X post URL."
        )

    username = match.group(1)
    post_id = match.group(2)

    return [{
        "url": (
            f"https://x.com/"
            f"{username}/status/"
            f"{post_id}"
        ),
        "handler_id": f"@{username}",
        "created_at": ""
    }]


# ============================================================
# GET OFFICIAL X EMBED HTML
# ============================================================

def get_x_oembed(post_url):

    oembed_url = "https://publish.x.com/oembed"

    params = {
        "url": post_url,
        "maxwidth": 550,
        "hide_thread": "true",
        "omit_script": "false",
        "lang": "en",
        "theme": "light"
    }

    print("Requesting official X oEmbed...")

    response = requests.get(
        oembed_url,
        params=params,
        timeout=30,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "(compatible; XReportGenerator/1.0)"
            )
        }
    )

    print(
        "X oEmbed STATUS:",
        response.status_code
    )

    if response.status_code != 200:
        print(
            "X oEmbed RESPONSE:",
            response.text[:1000]
        )

        raise Exception(
            "X oEmbed request failed: "
            f"{response.status_code}"
        )

    data = response.json()

    html = data.get("html")

    if not html:
        raise Exception(
            "X oEmbed returned no embed HTML."
        )

    print(
        "X oEmbed HTML received."
    )

    return html


# ============================================================
# CREATE SCREENSHOT PAGE
# ============================================================

def create_embed_page_html(embed_html):

    return f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">

    <meta
        name="viewport"
        content="width=device-width,
                 initial-scale=1.0"
    >

    <title>X Post</title>

    <style>

        html,
        body {{
            margin: 0;
            padding: 0;
            background: white;
        }}

        body {{
            width: 700px;
            min-height: 300px;

            display: flex;
            justify-content: center;
            align-items: flex-start;

            padding: 25px;
            box-sizing: border-box;
        }}

        #tweet-container {{
            width: 550px;
            max-width: 550px;
        }}

        blockquote.twitter-tweet {{
            margin: 0 auto !important;
        }}

    </style>
</head>

<body>

    <div id="tweet-container">
        {embed_html}
    </div>

    <script
        async
        src="https://platform.x.com/widgets.js"
        charset="utf-8">
    </script>

</body>
</html>
"""


# ============================================================
# WAIT FOR X EMBED TO RENDER
# ============================================================

def wait_for_x_embed(page):

    print(
        "Waiting for official X embed..."
    )

    # Wait for the original blockquote
    try:

        page.locator(
            "blockquote.twitter-tweet"
        ).wait_for(
            state="attached",
            timeout=15000
        )

        print(
            "X blockquote found."
        )

    except Exception:

        print(
            "X blockquote was not found."
        )

    # Wait for widgets.js rendering
    page.wait_for_timeout(5000)

    # X widget normally creates an iframe
    try:

        iframe = page.locator(
            "iframe"
        ).first

        iframe.wait_for(
            state="attached",
            timeout=20000
        )

        print(
            "X widget iframe found."
        )

        iframe.wait_for(
            state="visible",
            timeout=10000
        )

        print(
            "X widget iframe visible."
        )

    except Exception:

        print(
            "X widget iframe not found yet."
        )

    # Additional rendering time
    page.wait_for_timeout(5000)

    return True


# ============================================================
# TAKE REAL X EMBED SCREENSHOTS
# ============================================================

def take_screenshots(posts):

    successful_posts = []

    with sync_playwright() as p:

        print(
            "Starting Chromium..."
        )

        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu"
            ]
        )

        context = browser.new_context(
            viewport={
                "width": 800,
                "height": 1200
            },

            locale="en-US",

            timezone_id="UTC",

            color_scheme="light"
        )

        page = context.new_page()

        page.set_default_timeout(
            15000
        )

        for i, post in enumerate(
            posts,
            1
        ):

            print(
                f"Taking screenshot "
                f"{i}/{len(posts)}..."
            )

            try:

                post_url = post["url"]

                print(
                    "Post URL:",
                    post_url
                )

                # ------------------------------------------------
                # Get official X embed
                # ------------------------------------------------

                embed_html = get_x_oembed(
                    post_url
                )

                # ------------------------------------------------
                # Build local rendering page
                # ------------------------------------------------

                html = create_embed_page_html(
                    embed_html
                )

                print(
                    "Rendering official X embed..."
                )

                page.set_content(
                    html,
                    wait_until="domcontentloaded"
                )

                # ------------------------------------------------
                # Wait for widget
                # ------------------------------------------------

                wait_for_x_embed(
                    page
                )

                # ------------------------------------------------
                # Check rendered page
                # ------------------------------------------------

                print(
                    "Page title:",
                    page.title()
                )

                print(
                    "Iframe count:",
                    page.locator(
                        "iframe"
                    ).count()
                )

                # ------------------------------------------------
                # Find rendered X iframe
                # ------------------------------------------------

                iframe = None

                iframe_locator = page.locator(
                    "iframe"
                )

                for index in range(
                    iframe_locator.count()
                ):

                    candidate = (
                        iframe_locator
                        .nth(index)
                    )

                    try:

                        if (
                            candidate
                            .bounding_box()
                            is not None
                        ):

                            iframe = candidate
                            break

                    except Exception:
                        continue

                # ------------------------------------------------
                # Screenshot target
                # ------------------------------------------------

                filename = (
                    f"{SCREENSHOT_DIR}/"
                    f"post_{i}.png"
                )

                if iframe is not None:

                    print(
                        "Taking screenshot "
                        "of rendered X widget..."
                    )

                    iframe.screenshot(
                        path=filename
                    )

                else:

                    print(
                        "X iframe not found."
                    )

                    # The blockquote is still an
                    # official X representation.
                    blockquote = page.locator(
                        "blockquote.twitter-tweet"
                    ).first

                    if (
                        blockquote.count() > 0
                    ):

                        print(
                            "Taking screenshot "
                            "of X embed blockquote..."
                        )

                        blockquote.screenshot(
                            path=filename
                        )

                    else:

                        raise Exception(
                            "Official X embed "
                            "could not be rendered."
                        )

                if not os.path.exists(
                    filename
                ):

                    raise Exception(
                        "Screenshot file "
                        "was not created."
                    )

                post["screenshot"] = filename

                successful_posts.append(
                    post
                )

                print(
                    f"Screenshot {i} saved."
                )

            except Exception as e:

                print(
                    f"Screenshot {i} failed:"
                )

                print(
                    str(e)
                )

                traceback.print_exc()

                # Save debug screenshot
                try:

                    debug_file = (
                        f"{SCREENSHOT_DIR}/"
                        f"debug_{i}.png"
                    )

                    page.screenshot(
                        path=debug_file,
                        full_page=True
                    )

                except Exception:
                    pass

        context.close()
        browser.close()

    return successful_posts


# ============================================================
# CREATE WORD REPORT
# ============================================================

def create_word_report(posts):

    document = Document()

    title = document.add_heading(
        "X/Twitter Report",
        level=1
    )

    title.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    table = document.add_table(
        rows=1,
        cols=4
    )

    table.style = "Table Grid"

    headers = [
        "Sr.No",
        "Screenshot",
        "Post URL",
        "Handler ID"
    ]

    for i, header in enumerate(
        headers
    ):

        cell = (
            table.rows[0]
            .cells[i]
        )

        cell.text = header

        cell.vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )

        for paragraph in cell.paragraphs:

            paragraph.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            for run in paragraph.runs:

                run.bold = True

    for number, post in enumerate(
        posts,
        1
    ):

        cells = (
            table
            .add_row()
            .cells
        )

        cells[0].text = str(number)

        paragraph = (
            cells[1]
            .paragraphs[0]
        )

        paragraph.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
        )

        run = paragraph.add_run()

        run.add_picture(
            post["screenshot"],
            width=Inches(2.3)
        )

        cells[2].text = post["url"]

        cells[3].text = post["handler_id"]

        for cell in cells:

            cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.CENTER
            )

    output_file = (
        f"{REPORT_DIR}/X_Report.docx"
    )

    document.save(
        output_file
    )

    print(
        "Report saved:",
        output_file
    )

    return output_file


# ============================================================
# HOME PAGE
# ============================================================

@app.route("/")
def home():

    return render_template(
        "index.html"
    )


# ============================================================
# GENERATE REPORT
# ============================================================

@app.route(
    "/generate",
    methods=["POST"]
)
def generate():

    try:

        data = request.get_json(
            silent=True
        )

        if not data:

            return jsonify({
                "error":
                    "Invalid request."
            }), 400

        user_input = data.get(
            "input",
            ""
        ).strip()

        input_type = data.get(
            "input_type",
            "search"
        )

        number_of_posts = int(
            data.get(
                "number_of_posts",
                10
            )
        )

        if not user_input:

            return jsonify({
                "error":
                    "Please enter something."
            }), 400

        if input_type == "search":

            number_of_posts = max(
                1,
                min(
                    number_of_posts,
                    100
                )
            )

        print("")
        print(
            "=============================="
        )

        print(
            "NEW REPORT REQUEST"
        )

        print(
            "=============================="
        )

        # ------------------------------------------------------
        # SEARCH OR DIRECT URL
        # ------------------------------------------------------

        if input_type == "url":

            print(
                "Processing direct X URL..."
            )

            posts = get_post_from_url(
                user_input
            )

        else:

            print(
                "Searching X..."
            )

            print(
                "Requested posts:",
                number_of_posts
            )

            posts = search_posts(
                user_input,
                number_of_posts
            )

        print(
            "Posts found:",
            len(posts)
        )

        if not posts:

            return jsonify({
                "error":
                    "No posts found."
            }), 404

        # ------------------------------------------------------
        # SCREENSHOTS
        # ------------------------------------------------------

        print(
            "Starting screenshot process..."
        )

        posts = take_screenshots(
            posts
        )

        print(
            "Successful screenshots:",
            len(posts)
        )

        if not posts:

            return jsonify({
                "error":
                    "Could not capture the X posts."
            }), 500

        # ------------------------------------------------------
        # WORD REPORT
        # ------------------------------------------------------

        print(
            "Creating Word report..."
        )

        create_word_report(
            posts
        )

        print(
            "REPORT COMPLETE"
        )

        print(
            "=============================="
        )

        return jsonify({

            "success": True,

            "download_url":
                "/download",

            "posts":
                len(posts)

        })

    except Exception as e:

        print("")
        print(
            "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        )

        print(
            "ERROR WHILE GENERATING REPORT"
        )

        print(
            "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        )

        print(
            "ERROR:",
            str(e)
        )

        traceback.print_exc()

        print(
            "!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        )

        print("")

        return jsonify({
            "error":
                "Something went wrong."
        }), 500


# ============================================================
# DOWNLOAD REPORT
# ============================================================

@app.route("/download")
def download():

    file_path = (
        f"{REPORT_DIR}/X_Report.docx"
    )

    if not os.path.exists(
        file_path
    ):

        return jsonify({
            "error":
                "Report not found."
        }), 404

    return send_file(
        file_path,
        as_attachment=True,
        download_name="X_Report.docx"
    )


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":

    print("")
    print(
        "======================================"
    )

    print(
        "X/Twitter Report Generator"
    )

    print(
        "Starting Flask server..."
    )

    print(
        "======================================"
    )

    print("")

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=False,
        use_reloader=False
    )
```
