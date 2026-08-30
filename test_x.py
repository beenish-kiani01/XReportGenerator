import os
import requests
from dotenv import load_dotenv

load_dotenv()

token = os.getenv("X_BEARER_TOKEN")

url = "https://api.x.com/2/tweets/search/recent"

headers = {
    "Authorization": f"Bearer {token}"
}

params = {
    "query": "#AI -is:retweet",
    "max_results": 10,
    "tweet.fields": "created_at,author_id,text",
    "expansions": "author_id",
    "user.fields": "username"
}

response = requests.get(url, headers=headers, params=params)

print("Status:", response.status_code)
print(response.text)