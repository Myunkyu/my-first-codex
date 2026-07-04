from pathlib import Path
import re

src_path = Path('scripts/build_quantum_chapter2_methodology_docx.py')
source = src_path.read_text(encoding='utf-8')

# python-docx enum compatibility
source = source.replace(
    'from docx.enum.table import WD_ALIGN_VERTICAL, WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT',
    'from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT'
)

# Use wider printable area for wide analytical tables
source = source.replace('sec.left_margin = Cm(1.65)', 'sec.left_margin = Cm(1.20)')
source = source.replace('sec.right_margin = Cm(1.65)', 'sec.right_margin = Cm(1.20)')

# Matplotlib requires #-prefixed hex colors, while Word XML requires bare hex.
# Patch matplotlib-only keyword arguments before executing the source.
source = re.sub(r'(facecolor|edgecolor)=([A-Z_]+)', r'\1="#" + \2', source)
source = source.replace('facecolor=fc', 'facecolor="#" + fc')
source = source.replace('edgecolor=ec', 'edgecolor="#" + ec')
source = source.replace('facecolor=c', 'facecolor="#" + c')
source = source.replace('edgecolor=c', 'edgecolor="#" + c')
source = source.replace('color=GRAY', 'color="#" + GRAY')
source = source.replace('color=WHITE', 'color="#" + WHITE')
source = source.replace('color=NAVY', 'color="#" + NAVY')
source = source.replace('color=ec', 'color="#" + ec')
source = source.replace('color=c', 'color="#" + c')
source = source.replace('fc=ec', 'fc="#" + ec')
source = source.replace('ec=ec', 'ec="#" + ec')

# Avoid replacing function default arguments that belong to Word styling.
source = source.replace('def add_bullet(doc, text: str, level=0, color="#" + DARK,', 'def add_bullet(doc, text: str, level=0, color=DARK,')
source = source.replace('def add_body(doc, text: str, bold=False, color="#" + DARK,', 'def add_body(doc, text: str, bold=False, color=DARK,')

compiled = compile(source, str(src_path), 'exec')
namespace = {'__name__': '__main__', '__file__': str(src_path)}
exec(compiled, namespace)
