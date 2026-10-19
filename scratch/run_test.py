import urllib.request
import urllib.parse
import json
import time
from pathlib import Path

tex = Path("jobs/e703e503/document.tex").read_text(encoding="utf-8")
form_fields = {
    "latex_source": tex,
    "mode": "1",
    "dpi": "400",
    "strict": "false",
    "qa_pdf": "true",
}
data = urllib.parse.urlencode(form_fields).encode("utf-8")
req = urllib.request.Request("http://localhost:8000/api/convert/paste", data=data, method="POST")

t0 = time.time()
print("Starting conversion request...")
with urllib.request.urlopen(req, timeout=180) as resp:
    res = json.loads(resp.read().decode("utf-8"))

elapsed = round(time.time() - t0, 2)
print(f"Conversion finished in {elapsed}s!")
print(f"Job ID: {res.get('job_id')}")
print(f"Overall Status: {res.get('overall')}")
print("5-Layer Results:")
for lr in res.get("layers", []):
    print(f"  Layer {lr.get('layer')}: {lr.get('name')} -> [{lr.get('status')}] {lr.get('summary')}")
