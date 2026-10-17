"""Verify the official ECDICT blob and build bounded, on-demand lookup buckets.

Maintainer-only: installation and lookup do not require Python or network access.
"""
import base64
import csv
import hashlib
import itertools
import json
from pathlib import Path
import sqlite3

root = Path(__file__).resolve().parents[1]
metadata = json.loads((root / "bin/ecdict-remote.json").read_text(encoding="utf-8-sig"))
csv_path = root / "bin/ecdict-full.csv"
if not csv_path.exists():
    blob = json.loads((root / "bin/ecdict-blob.json").read_text(encoding="utf-8-sig"))
    raw = base64.b64decode(blob["content"])
    del blob
    assert len(raw) == metadata["size"], "Incomplete dictionary download"
    actual = hashlib.sha1(b"blob " + str(len(raw)).encode() + b"\0" + raw).hexdigest()
    assert actual == metadata["sha"], "Dictionary Git blob hash mismatch"
    csv_path.write_bytes(raw)
    del raw
else:
    raw = csv_path.read_bytes()
    assert hashlib.sha1(b"blob " + str(len(raw)).encode() + b"\0" + raw).hexdigest() == metadata["sha"]
    del raw

target = root / "data/full"
target.mkdir(exist_ok=True)
csv.field_size_limit(4 * 1024 * 1024)
total_rows = 0
with sqlite3.connect(root / "bin/full-lexicon-build.sqlite") as db:
    db.execute("drop table if exists entries")
    db.execute("create table entries(word text primary key,ipa text,meaning text,bucket text)")
    with csv_path.open(encoding="utf-8-sig", newline="") as source:
        for row in csv.DictReader(source):
            total_rows += 1
            word = row["word"].strip().lower()
            meaning = row.get("translation", "").replace("\\n", "\n").strip()
            if not word or not any("\u3400" <= char <= "\u9fff" for char in meaning):
                continue
            ipa = row.get("phonetic", "").strip().strip("/[]")
            if ipa:
                ipa = "/" + ipa.replace("ә", "ə").replace("є", "ɛ").replace("'", "ˈ").replace(":", "ː") + "/"
            bucket = word[:2].encode("utf-8").hex() + ".json"
            db.execute("insert or replace into entries values(?,?,?,?)", (word, ipa, meaning, bucket))
    db.commit()
    count = db.execute("select count(*) from entries").fetchone()[0]
    assert total_rows > 750000 and count > 700000, "Unexpectedly small dictionary"
    manifest = {"entries": count, "sourceRows": total_rows, "sourceGitBlob": metadata["sha"],
                "source": "https://github.com/skywind3000/ECDICT", "license": "MIT", "buckets": {}}
    cursor = db.execute("select bucket,word,ipa,meaning from entries order by bucket,word")
    for name, rows in itertools.groupby(cursor, key=lambda item: item[0]):
        entries = {word: [ipa, meaning] for _, word, ipa, meaning in rows}
        payload = json.dumps(entries, ensure_ascii=False, separators=(",", ":")).encode()
        temporary = target / (name + ".tmp")
        temporary.write_bytes(payload)
        temporary.replace(target / name)
        manifest["buckets"][name] = {"entries": len(entries), "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}
    (target / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print("Verified official Git blob:", metadata["sha"])
print("Source rows:", total_rows, "Chinese entries:", count, "buckets:", len(manifest["buckets"]))
print("Max bucket bytes:", max(item["bytes"] for item in manifest["buckets"].values()))
