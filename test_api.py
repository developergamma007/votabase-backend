import urllib.request
import json
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

req = urllib.request.Request("http://127.0.0.1:8000/votebase/v1/api/volunteers?page=0&size=5")
# We need to bypass auth or auth is disable locally?
# Actually in main.py, list_volunteers uses "Depends(require_roles(...)". We need a token.
