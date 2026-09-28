import json
import os

log_path = r"C:\Users\zillur\.gemini\antigravity-ide\brain\3c16a7bd-b6ce-4539-b3a8-7e1f8e88dba8\.system_generated\logs\transcript_full.jsonl"
target_tsv = "client_cash_book_data.tsv"

with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
    for line in f:
        try:
            data = json.loads(line)
            if data.get("type") == "USER_INPUT":
                content = data.get("content", "")
                if "Receipts" in content and "Premier Bank" in content:
                    with open(target_tsv, "w", encoding="utf-8") as out:
                        out.write(content)
                    print(f"Saved {len(content)} chars to {target_tsv}")
                    break
        except Exception:
            pass
