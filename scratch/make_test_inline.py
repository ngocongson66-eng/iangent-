import json

with open('scratch/script_0.js', 'r', encoding='utf-8') as f:
    js_code = f.read()

# Replace all 4 placeholders exactly as server.py does
js_code = js_code.replace('__INITIAL_SAMPLE_JSON__', '[]')
js_code = js_code.replace('__INITIAL_EXTRACTED_JSON__', '[]')
js_code = js_code.replace('__INITIAL_DOSSIERS_JSON__', '[]')
js_code = js_code.replace('__INITIAL_SUMMARY_JSON__', '{}')

html = f"""<!DOCTYPE html>
<html>
<head>
<script>
window.__errors = [];
window.onerror = function(msg, url, line, col, error) {{
    window.__errors.push({{msg: msg, line: line, col: col, error: String(error)}});
}};
</script>
</head>
<body>
<div id="test-output">Testing inline JS...</div>
<script>
{js_code}
</script>
<script>
var div = document.getElementById('test-output');
if (window.__errors.length > 0) {{
    div.innerText = 'EXACT_ERROR: ' + JSON.stringify(window.__errors);
}} else {{
    div.innerText = 'SUCCESS: NO SYNTAX ERRORS! typeof switchTask = ' + typeof switchTask;
}}
</script>
</body>
</html>
"""

with open('scratch/test_inline.html', 'w', encoding='utf-8') as f:
    f.write(html)
print("Updated scratch/test_inline.html with all 4 placeholders")
