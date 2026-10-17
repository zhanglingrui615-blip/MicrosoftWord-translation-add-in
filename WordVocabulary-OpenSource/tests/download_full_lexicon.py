"""Resumable official ECDICT download. Verify the Git blob before publishing CSV."""
from concurrent.futures import ThreadPoolExecutor, as_completed
import hashlib
import json
from pathlib import Path
import time
import urllib.request

root = Path(__file__).resolve().parents[1] / "bin"
metadata = json.loads((root / "ecdict-remote.json").read_text(encoding="utf-8-sig"))
size, sha = metadata["size"], metadata["sha"]
chunk_size = 1024 * 1024
parts = root / "ecdict-parts"
parts.mkdir(exist_ok=True)
count = (size + chunk_size - 1) // chunk_size

# Reuse only complete chunks of the already downloaded raw prefix; final SHA validates them too.
prefix = root / "ecdict-full.download"
if prefix.exists():
    with prefix.open("rb") as existing:
        if existing.read(5) == b"word,":
            existing.seek(0)
            for i in range(prefix.stat().st_size // chunk_size):
                part = parts / f"{i:03d}.part"
                data = existing.read(chunk_size)
                if not part.exists():
                    part.write_bytes(data)

def download(i):
    begin = i * chunk_size
    length = min(chunk_size, size - begin)
    path = parts / f"{i:03d}.part"
    if path.exists() and path.stat().st_size == length:
        return i
    last = None
    for attempt in range(5):
        url = ("https://raw.githubusercontent.com/skywind3000/ECDICT/master/ecdict.csv"
               if attempt % 2 == 0 else f"https://api.github.com/repos/skywind3000/ECDICT/git/blobs/{sha}")
        headers = {"Range": f"bytes={begin}-{begin + length - 1}", "Accept-Encoding": "identity",
                   "Accept": "application/vnd.github.raw+json", "User-Agent": "WordVocabulary"}
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=40) as response:
                if response.status != 206 or not response.headers.get("Content-Range", "").startswith(f"bytes {begin}-{begin + length - 1}/"):
                    raise ValueError("Server returned an unexpected range")
                data = response.read(length + 1)
            if len(data) != length:
                raise ValueError("Incomplete chunk")
            temp = parts / f"{i:03d}.tmp"
            temp.write_bytes(data)
            temp.replace(path)
            return i
        except Exception as error:
            last = error
            time.sleep(min(1 + attempt, 4))
    raise RuntimeError(f"Failed to download chunk {i}: {last}")

completed = 0
with ThreadPoolExecutor(max_workers=8) as pool:
    for future in as_completed([pool.submit(download, i) for i in range(count)]):
        future.result()
        completed += 1
        if completed % 5 == 0 or completed == count:
            print(f"Verified-size chunks: {completed}/{count}", flush=True)

digest = hashlib.sha1(f"blob {size}\0".encode())
target = root / "ecdict-full.csv.tmp"
with target.open("wb") as output:
    for i in range(count):
        data = (parts / f"{i:03d}.part").read_bytes()
        digest.update(data)
        output.write(data)
assert target.stat().st_size == size and digest.hexdigest() == sha, "Official Git blob integrity check failed"
target.replace(root / "ecdict-full.csv")
print(f"Complete official dictionary verified: {size} bytes, Git blob {sha}", flush=True)
