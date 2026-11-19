# builds app.html from src/ (the reading engine is placed inside the page: one file, no extra loading)
import pathlib
root = pathlib.Path(__file__).parent
html = (root / 'src/app.html').read_text(encoding='utf-8')
core = (root / 'src/core.js').read_text(encoding='utf-8')
(root / 'app.html').write_text(html.replace('/*CORE*/', core), encoding='utf-8')
print('app.html built')
