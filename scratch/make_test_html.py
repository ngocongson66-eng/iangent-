# Create a test HTML file that includes scratch/script_0.js and captures any syntax error!
html_content = """<!DOCTYPE html>
<html>
<head>
<script>
window.__errors = [];
window.onerror = function(msg, url, line, col, error) {
    window.__errors.push({msg: msg, line: line, col: col, error: String(error)});
};
</script>
</head>
<body>
<div id="test-output">Testing JS...</div>
<script src="script_0.js"></script>
<script>
var div = document.getElementById('test-output');
if (window.__errors.length > 0) {
    div.innerText = 'ERRORS: ' + JSON.stringify(window.__errors);
} else {
    div.innerText = 'SUCCESS: NO SYNTAX ERRORS! typeof switchTask = ' + typeof switchTask;
}
</script>
</body>
</html>
"""
with open('scratch/test_syntax.html', 'w', encoding='utf-8') as f:
    f.write(html_content)
print("Created scratch/test_syntax.html")
