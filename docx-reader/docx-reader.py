import os
from zipfile import ZipFile

os.chdir('')

from bs4 import BeautifulSoup as BS

def read_docx(filename):
    doc = ZipFile(filename)
    content_xml = doc.read('word/document.xml')
    return BS(content_xml, 'xml')

def find_iter(soup, tagname):
    tag = soup.find(tagname)
    while tag is not None:
        yield tag
        tag = tag.find_next(tagname)
        
def iter_paragraphs(soup):
    return find_iter(soup, 'w:p')

def try_or_none_properties(properties, arg1, tag):
    try:
        return properties.find(arg1).get(tag)
    except AttributeError:
        try:
            properties.find(arg1)
            return '1'
        except AttributeError:
            return None

doc = read_docx('TESTDOC.docx')

find = False
for i, paragraph in enumerate(iter_paragraphs(doc)):
    paragraph
    if paragraph.text == 'Simple liste lvl 1':
        break
    
    properties = paragraph.find('w:pPr')
    
    style = try_or_none_properties(properties, 'w:pStyle', 'w:val')
    list_lvl = try_or_none_properties(properties, 'w:ilvl', 'w:val')
    list_type = try_or_none_properties(properties, 'w:numId', 'w:val')
    position = try_or_none_properties(properties, 'jc', 'w:val')
#    try_or_none_properties(properties, 'ind', 'w:firstLineChars')
#    try_or_none_properties(properties, 'ind', 'w:hanging')
#    try_or_none_properties(properties, 'ind', 'w:left')
#    try_or_none_properties(properties, 'ind', 'w:leftChars')
    
    print(style, list_lvl, list_type)
    
    runs_prop = []
    run_prop = ['b', 'bCs','i', 'iCs', 'color', 'highlight', 'vertAlign', 'u']
    try:
        for run in paragraph.find_all('r'):
            prop = run.find('rPr')
            runs_prop.append([try_or_none_properties(prop, x, 'w:val') for x in run_prop])
    except AttributeError:
        pass
    print(runs_prop)
    
    tabs_prop = []
    tabs = properties.find('w:tabs')
    for i in tabs.find_all('w:tab'):
        tabs_prop.append([i.get('w:pos'), i.get('w:val')])

    
