import re
with open(r'E:\Construction APP\Files\attendance_data.js', 'r', encoding='utf-8-sig') as f:
    content = f.read()
m = re.search(r'const\s+ATTENDANCE_DATA\s*=\s*(\[.*?\]);', content, re.DOTALL)
print(f"Matched: {m is not None}")
print(f"File length: {len(content)}")
if m:
    raw = m.group(1)
    print(f"Array length: {len(raw)}")
    # Show last 80 chars, replacing non-ascii with ?
    safe = raw[-80:].encode('ascii', 'replace').decode()
    print(f"Last 80 chars: {safe}")
else:
    safe = content[-100:].encode('ascii', 'replace').decode()
    print(f"Last 100 chars: {safe}")
