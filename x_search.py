import os
import requests
from dotenv import load_dotenv

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

    if response.status_code != 200:
        print("API Error:", response.status_code)
        print(response.text)
        return []

    result = response.json()

    users = {
        user["id"]: user["username"]
        for user in result.get("includes", {}).get("users", [])
    }

    posts = []

    for tweet in result.get("data", []):
        author_id = tweet.get("author_id")
        username = users.get(author_id, "Unknown")

        posts.append({
            "id": tweet["id"],
            "url": f"https://x.com/{username}/status/{tweet['id']}",
            "handler_id": f"@{username}",
            "text": tweet.get("text", "")
        })

    return posts[:number_of_posts]


if __name__ == "__main__":
    query = input("Enter keyword or hashtag: ")
    number = int(input("How many posts? "))

    posts = search_posts(query, number)

    print(f"\nFound {len(posts)} posts:\n")

    for i, post in enumerate(posts, 1):
        print(f"{i}. {post['handler_id']}")
        print(f"   {post['url']}")
        print(f"   {post['text'][:100]}")
        print()