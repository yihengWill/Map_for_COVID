import requests
import json

# To crawler the json from the website
url = 'https://view.inews.qq.com/g2/getOnsInfo?name=disease_h5'
headers = {
'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
              'Connection': 'keep-alive',
              "Referer": "https://news.qq.com"
}
response = requests.get(url,headers=headers).json()

# To write the json for China.json, because using Chinese so ensure_ascii=False
json = json.loads(response['data'])
with open('./China.json', 'w', encoding="utf-8") as f:
    f.write(json.dumps(json, ensure_ascii=False, indent=2))
