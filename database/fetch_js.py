import urllib.request
import os

url = "http://localhost:3000/app.js?v=555"
dest = r"C:\Users\jayav\.gemini\antigravity-ide\brain\b0067744-fd02-4ccf-8e27-b9dca0c0bba3\scratch\served_app.js"

try:
    print(f"Fetching {url}...")
    with urllib.request.urlopen(url, timeout=5) as response:
        content = response.read()
        with open(dest, "wb") as f:
            f.write(content)
    print(f"Success! Saved to {dest} (Size: {os.path.getsize(dest)} bytes)")
except Exception as e:
    print(f"Error fetching: {e}")
