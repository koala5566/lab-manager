"""把 tools/webapp/ 的原始檔組成 apps-script/16_網頁工作台.gs。

用法：python3 tools/webapp/build_webapp.py
畫面要改就改 app.html / app.css / app.js / art.js，再執行這支程式。
"""
import os

here = os.path.dirname(os.path.abspath(__file__))
read = lambda n: open(os.path.join(here, n), encoding='utf8').read()
html = read('app.html').replace('/*CSS*/', read('app.css')).replace('/*JS*/', read('art.js') + '\n' + read('app.js'))
# 放進 JavaScript 的 `...` 字串：跳脫反斜線、反引號、${
lit = html.replace('\\', '\\\\').replace('`', '\\`').replace('${', '\\${')
out = read('server.gs') + '\n// ---------------------------------------------------------------- 畫面（自動產生，請改 tools/webapp/ 的原始檔）\n\nconst APP_HTML = `' + lit + '`;\n'
dst = os.path.join(here, '..', '..', 'apps-script', '16_網頁工作台.gs')
open(dst, 'w', encoding='utf8').write(out)
print('寫入', os.path.normpath(dst), len(out.splitlines()), '行')
