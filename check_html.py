with open(r'E:\Construction APP\Marks Regester\index.html', 'r', encoding='utf-8') as f:
    content = f.read()
print('File OK, length:', len(content))
print('Has window.onload:', 'window.onload' in content)
print('Has init function:', 'function init()' in content)
print('Has GROUPS_DATA:', 'const GROUPS_DATA' in content)
# Find the script section
start = content.find('<script>')
end = content.find('</script>')
js = content[start+8:end]
print('JS length:', len(js))
# Check for unclosed strings in JS
lines = js.split('\n')
for i, line in enumerate(lines):
    stripped = line.strip()
    if stripped.startswith('//'): continue
    single_count = stripped.count("'")
    double_count = stripped.count('"')
    if single_count % 2 != 0:
        print(f"  Odd single quotes at line {i+1}: {stripped[:80]}")
    if double_count % 2 != 0:
        print(f"  Odd double quotes at line {i+1}: {stripped[:80]}")
print('Check complete')
