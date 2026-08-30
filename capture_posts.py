import os
import requests
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()

TOKEN = os.getenv("X_BEARER_TOKEN")


def search_posts(query, number_of_posts):
    url = "https://api.x.com/2/tweets/search/recent"

    headers = {
        "Authorization": f"Bearer {TOKEN}"
    }

    params = {
        "query": query + " -is:retweet",
        "max_results": max(10, min(number_of_posts, 100)),
        "tweet.fields": "created_at,author_id",
        "expansions": "author_id",
        "user.fields": "username"
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    result = response.json()

    users = {
        user["id"]: user["username"]
        for user in result.get("includes", {}).get("users", [])
    }

    posts = []

    for tweet in result.get("data", []):
        username = users.get(tweet.get("author_id"), "Unknown")

        posts.append({
            "url": f"https://x.com/{username}/status/{tweet['id']}",
            "handler_id": f"@{username}"
        })

    return posts[:number_of_posts]


def take_screenshots(posts):
    os.makedirs("screenshots", exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(
            viewport={"width": 1400, "height": 1200}
        )

        for i, post in enumerate(posts, 1):
            print(f"Taking screenshot {i}/{len(posts)}...")

            page.goto(post["url"], wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(5000)

            filename = f"screenshots/post_{i}.png"
            page.screenshot(path=filename, full_page=False)

            post["screenshot"] = filename

        browser.close()


if __name__ == "__main__":
    query = input("Enter keyword or hashtag: ")
    number = int(input("How many posts? "))

    posts = search_posts(query, number)

    print(f"\nFound {len(posts)} posts.")

    take_screenshots(posts)

with open("posts.json", "w", encoding="utf-8") as file:
    import json
    json.dump(posts, file, indent=4)

print("\nDone!")
print("Post information saved to posts.json")