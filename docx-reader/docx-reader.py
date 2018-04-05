import os
from zipfile import ZipFile
import bs4
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup as BS

os.chdir('')


def read_docx(filename):
    # Read the docx document and return a BeautifulSoup object
    doc = ZipFile(filename)
    content_xml = doc.read('word/document.xml')
    return BS(content_xml, 'xml')

def find_iter(soup, tagname):
    # Generator function to iter through elements of a document efficiently
    # Does not work for a subitem of a document's soup, because it will find
    # values outside its scope
    tag = soup.find(tagname)
    while tag is not None:
        yield tag
        tag = tag.find_next(tagname)
        
def iter_paragraphs(soup):
    # Simple use of the find_iter function with predefined tagname
    return find_iter(soup, 'w:p')

def try_or_none_properties(properties, arg1, tag):
    try:
        finder = properties.find(arg1).get(tag)
        if not finder:
            return True
        else:
            return finder
    except AttributeError:
        try:
            properties.find(arg1).name
            return True
        except AttributeError:
            return None
        
def document_summary(doc):
    '''
    Input : document name
    Output : pandas dataframe with characteristics (not pretty nor complete yet)
    '''
    doc = read_docx(doc)
    summary = list()
    find = False
    for i, paragraph in enumerate(iter_paragraphs(doc)):
        text = paragraph.text

        properties = paragraph.find('w:pPr')

        style = try_or_none_properties(properties, 'w:pStyle', 'w:val')
        list_lvl = try_or_none_properties(properties, 'w:ilvl', 'w:val')
        list_type = try_or_none_properties(properties, 'w:numId', 'w:val')
        position = try_or_none_properties(properties, 'jc', 'w:val')
    #    try_or_none_properties(properties, 'ind', 'w:firstLineChars')
    #    try_or_none_properties(properties, 'ind', 'w:hanging')
    #    try_or_none_properties(properties, 'ind', 'w:left')
    #    try_or_none_properties(properties, 'ind', 'w:leftChars')
        
        # the output doesn't work yet with multiple run per paragraph
        runs_prop = []
        run_prop = ['b', 'bCs','i', 'iCs', 'color', 'highlight', 'vertAlign', 'u']
        try:
            for run in find_iter(paragraph, 'r'):
                prop = run.find('rPr')
                runs_output = [try_or_none_properties(prop, x, 'w:val') for x in run_prop]
                runs_prop.append(runs_output)
            runs_prop = [x[0] if len(set(x)) == 1 else x for x in np.array(runs_prop).T]
        except AttributeError:
            runs_prop = [None] * 8

        tabs_prop = []
        try: 
            tabs = properties.find('w:tabs')
            for i in tabs.find_all('w:tab'):
                tabs_prop.append([i.get('w:pos'), i.get('w:val')])
        except AttributeError:
            pass

        characteristics = ['text', 'paragraph_style', 'list_lvl','list_type' ,'horizontal_alignment', 'vertical_alignment',\
                           'bold', 'italic', 'text_color', 'highlight', 'underline']

        try:
            summary.append([text, style, list_lvl, list_type, position, runs_prop[6], runs_prop[0], runs_prop[2],\
                        runs_prop[4], runs_prop[5], runs_prop[7]])
        except IndexError: # because of issue that doesn't allow to create a list of None in a loop
            summary.append([text, style, list_lvl, list_type, position, None, None, None,\
                        None, None, None])
            
    return pd.DataFrame(summary, columns = characteristics)
