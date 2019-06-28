import sys
import re
from pathlib import Path
import pypandoc
import pptx

def main(path : Path):
    if not path.exists(): sys.exit()
    with path.open('r') as inf:
        input_md = inf.read()
    blank_md = input_md

    for match in re.finditer(r'\*\*(.+?)\*\*', input_md):
        length = match.regs[1][1] - match.regs[1][0]
        rep = '__' * length
        strmatch = match.string[match.regs[0][0]:match.regs[0][1]]
        blank_md = blank_md.replace(strmatch, f'{rep}')
    pypandoc.convert_text(blank_md, 'pdf', format = 'gfm',
        outputfile=str(path.parent / (path.stem + '.pdf')),
        extra_args=[
            '--pdf-engine=wkhtmltopdf',
            '--to=html5',
            '--css=style.css',
            '-V', 'margin-left=1in',
            '-V', 'margin-right=1in',
            '-V', 'margin-top=1in',
            '-V', 'margin-bottom=1in'])
    
    input_pptx = input_md.splitlines()
    del input_pptx[2:]
    for match in re.finditer(r'[0-9]+\. (.+?)(?=([0-9]+\.)|\Z)', input_md, re.MULTILINE | re.DOTALL):
        input_pptx.append('# ' + input_md[match.regs[1][0]:match.regs[1][1]].strip())
    pypandoc.convert_text('\n'.join(input_pptx), 'pptx', format = 'gfm',
        outputfile=str(path.parent / (path.stem + '.pptx')),
        extra_args=[
            '--reference-doc=custom-reference.pptx'
        ])

if __name__ == "__main__":
    main(Path(sys.argv[1]))