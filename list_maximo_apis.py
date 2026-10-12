import requests
import urllib3
import json
from dotenv import load_dotenv
import os 

load_dotenv()

urllib3.disable_warnings()

API_KEY = os.getenv("MAXIMO_API_KEY")

url = (
    "https://web.manage.demo.apps.mas9.eam-commit.com"
    "/maximo/api/apimeta"
)

response = requests.get(
    url,
    headers={
        "apikey": API_KEY,
        "Accept": "application/json"
    },
    verify=False
)

data = response.json()

print(json.dumps(data, indent=2))