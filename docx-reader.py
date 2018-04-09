import os
from zipfile import ZipFile
import bs4
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup as BS

os.chdir(')

class docx_reader:
    def __init__(self, filename):
        self.filename = filename
        self.doc = self._read_docx()
        
        summary = list()
        tables = self.doc.find_all('tbl')
        
        # change the way to indicate if it's in table: detect if it's wrapped in a table
        # and then add index based on the total length of the tables
        # also add the cell idnex based on w:tblW w:type="dxa" w:w="8522"
        # and w:gridCol w:w="4261"
        
        for i, paragraph in enumerate(self._iter_paragraphs()):
            index_table = None
            
            text = paragraph.text
    
            properties = paragraph.find('w:pPr')
    
            style = self._try_or_none_properties(properties, 'w:pStyle', 'w:val')
            list_lvl = self._try_or_none_properties(properties, 'w:ilvl', 'w:val')
            list_type = self._try_or_none_properties(properties, 'w:numId', 'w:val')
            position = self._try_or_none_properties(properties, 'jc', 'w:val')
            
            runs_prop = []
            run_prop = [ 'vertAlign', 'b', 'i', 'color', 'highlight', 'u', 'bCs', 'iCs']
            try:
                for run in paragraph.find_all('r'):
                    prop = run.find('rPr')
                    runs_output = [self._try_or_none_properties(prop, x, 'w:val') for x in run_prop]
                    runs_prop.append(runs_output)
                runs_prop = [x[0] if len(set(x)) == 1 else x for x in np.array(runs_prop).T]
            except AttributeError:
                runs_prop = [None] * 8
    
#            tabs_prop = []
#            try: 
#                tabs = properties.find('w:tabs')
#                for i in tabs.find_all('w:tab'):
#                    tabs_prop.append([i.get('w:pos'), i.get('w:val')])
#            except AttributeError:
#                pass
            
            bookmark = self._try_or_none_properties(paragraph, 'bookmarkStart', 'w:id')
            
            for i, table in enumerate(tables):
                if paragraph in table.find_all('p'):
                    index_table = 'table_' + str(i)
            
            try : img_id = "_".join(['image', paragraph.find('cNvPr').get('id')])
            except: img_id = None
    
            characteristics = ['text', 'paragraph_style', 'list_lvl','list_type' ,\
                               'horizontal_alignment', 'vertical_alignment',\
                               'bold', 'italic', 'text_color', 'highlight', 'underline', \
                               'bookmark', 'table', 'cell', 'image']
    
            try:
                summary.append([text, style, list_lvl, list_type, position,\
                                runs_prop[:-2], bookmark, index_table, \
                                cell_index, img_id])
                
            except IndexError: # because of issue that doesn't allow to create a list of None in a loop
                summary.append([text, style, list_lvl, list_type, position, \
                                None, None, None, None, None, None, bookmark, \
                                index_table, cell_index, img_id])
        
        output = pd.DataFrame(summary, columns = characteristics)
        
        output[output == '0'] = None
        output[output == 'none'] = None
        
        output.text_color[output.text_color == 'auto'] = None
        
        
        self.summary = output
            
        
    def _read_docx(self):
        # Read the docx document and return a BeautifulSoup object
        doc = ZipFile(self.filename)
        content_xml = doc.read('word/document.xml')
        return BS(content_xml, 'xml')
    
    def _find_iter(self, tagname):
        # Generator function to iter through elements of a document efficiently
        # Does not work for a subitem of a document's soup, because it will find
        # values outside its scope
        tag = self.doc.find(tagname)
        while tag is not None:
            yield tag
            tag = tag.find_next(tagname)
            
    def _iter_paragraphs(self):
        # Simple use of the find_iter function with predefined tagname
        return self._find_iter('w:p')
    
    def _try_or_none_properties(self, properties, arg1, tag):
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

filename = 'TESTDOC.docx'

doc = docx_reader(filename)
