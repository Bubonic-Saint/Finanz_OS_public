import http.client
import os
from dotenv import load_dotenv


load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', 'data', '.env'))


def get_transactions(date_from: str = None):
    base_url = os.getenv("BASE_URL").removeprefix("https://").removeprefix("http://")
    conn = http.client.HTTPSConnection(base_url)

    headers = {'x-api-key': os.getenv("API_KEY")}

    endpoint = os.getenv("GET_TRANSACTIONS").format(accountId=os.getenv("ACCOUNT_ID"))
    conn.request("GET", endpoint, headers=headers)

    res = conn.getresponse()
    import json
    data = json.loads(res.read().decode("utf-8"))

    if date_from:
        data["transactions"] = [
            t for t in data["transactions"] if t["date"] >= date_from
        ]

    return data
