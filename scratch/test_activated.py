import requests, json, os

base = 'http://127.0.0.1:8000'
url = f"{base}/votebase/v1/api/message-template/activated-wards?assemblyId=170"
try:
    resp = requests.get(url, timeout=5)
    print('Status:', resp.status_code)
    print('Body:', resp.text)
except Exception as e:
    print('Error calling API:', e)
