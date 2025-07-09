import zipfile, xml.etree.ElementTree as ET, sys, os

INPUT_DOCX='DRAFT_THESE_ VF_26 06 25 .docx'
OUTPUT_TEX='DRAFT_THESE_ VF_26 06 25 .tex'

# Extract document.xml
with zipfile.ZipFile(INPUT_DOCX) as z:
    with z.open('word/document.xml') as f:
        xml_data=f.read()

ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
root=ET.fromstring(xml_data)

def paragraph_to_tex(style,text):
    tex=text.replace('&','\&').replace('%','\%').replace('_','\_')
    if style in ('Titre1','Chapitre'):
        return f"\\chapter{{{tex}}}\n"
    elif style=='Titre2':
        return f"\\section{{{tex}}}\n"
    elif style=='Titre3':
        return f"\\subsection{{{tex}}}\n"
    else:
        return tex+'\n\n'

tex_content=[]
for p in root.findall('.//w:p', ns):
    text=''.join(t.text or '' for t in p.findall('.//w:t', ns)).strip()
    if not text:
        continue
    pPr=p.find('w:pPr', ns)
    style=None
    if pPr is not None:
        pStyle=pPr.find('w:pStyle', ns)
        if pStyle is not None:
            style=pStyle.get(f'{{{ns["w"]}}}val')
    tex_content.append(paragraph_to_tex(style,text))

header=open('header.tex').read()
with open(OUTPUT_TEX,'w',encoding='utf-8') as f:
    f.write(header)
    f.write('\n')
    for t in tex_content:
        f.write(t)
    f.write('\n\\end{document}\n')
