#import os
from zipfile import ZipFile
#import bs4
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup as BS

#os.chdir('')

class docx_reader:
    def __init__(self, filename):
        self.filename = filename
        self.doc = self._read_docx()
        
        summary = list()
        
        # find all the tables in the document for later indexing
        self._find_tables_composition()
        
        
        # initial values for table/cell indexing
        # currently, only tables, subtables (table inside a table) and sub-subtables
        # are taken into account
        current_table = 0
        current_cell = 0
        current_cell_index = [0,0]
        current_paragraph = 0
        current_table_width = int(tables[current_table].find('tblW').get('w:w'))
        current_table_nb_cells = len(tables[current_table].find_all('tc'))
        current_nb_paragraph_in_cell = [len(x.find_all('p')) for x in \
                                        tables[current_table].find_all('tc')]
        previous_table_lvl = 0
        tables_indices = [0 for j in range(len(self.tables_colnames))]
#        paragraphs_indices = self.p_per_tables.copy()
#        table_index = 0
#        tables_text = [None for j in range(len(self.tables_colnames))]
        p_indices = [x for i, x in enumerate(self.p_per_tables) if i not in self.index_subtables]
        
        for i, paragraph in enumerate(self._iter_paragraphs()):
#            index_table = None
#            cell_index = None
            
            # if we have a table as parent, then ...
            up_table = paragraph.find_parent('tbl')
            if not up_table:
                cell_index = None
            
            else:
#                while True:
#                    if paragraphs_indices[table_index] == 0:
#                        table_index += 1
##                        if previous_table_lvl == 0:
##                            tables_indices[previous_table_lvl] +=1
#                    else:
#                        break
                if p_indices[tables_indices[0]] == 0:
                    tables_indices[0] += 1
                
                p_indices[tables_indices[0]] -= 1
                table_lvl = 0
                a = up_table
                while a.find_parent('tbl'):
                    table_lvl += 1
                    a = a.find_parent('tbl')
                
#                print(table_lvl)
                if table_lvl > previous_table_lvl:
#                    print(paragraph.text)
#                    print('up {} lvl'.format(str(table_lvl-previous_table_lvl)))
                    tables_indices[table_lvl] += 1 
                    previous_table_lvl = table_lvl
                elif table_lvl < previous_table_lvl:
#                    print(paragraph.text)
#                    print('dn {} lvl'.format(str(-table_lvl+previous_table_lvl)))
                    previous_table_lvl = table_lvl
                
                tables_text = []
                for i in range(len(self.tables_colnames)):
                    try:
                        tables_text.append(str(tables_indices[:table_lvl+1][i]))
                    except IndexError:
                        tables_text.append(None)
                        
                print(tables_text)
                
#                print(tables_indices[:table_lvl+1])
                
                
#                paragraphs_indices[table_index] -= 1
                
#                print(table_index)
#                if paragraphs_indices[table_index] == 0 and table_lvl == 0:
#                    tables_indices[0] +=1
#                
#                tableIndex = [i for i, x in enumerate(self.tables) if str(x) == str(up_table)]
                
                
#                print(tableIndex)
            
                # values to be outputed
#                cell_index = "_".join([str(int(x)) for x in current_cell_index])
#                index_table = str(current_table + 1)
#                
#                # we are looking at a new paragraph inside the cell
#                current_paragraph += 1
#                                
#                # we now have to define if the new cell finished, if so we 
#                # add its width/table_width to see if we start a new row or not
#                if current_paragraph == current_nb_paragraph_in_cell[current_cell]:
#                    add_width = int(paragraph.find_parent('tc').find('tcW')\
#                                    .get('w:w'))/current_table_width
#                    if int(current_cell_index[0]) !=\
#                                int(current_cell_index[0] + add_width):
#                        current_cell_index[1] = 0
#                    else :
#                        current_cell_index[1] += 1
#                    current_cell_index[0] += add_width
#                    
#                    # then we increase the cell count and reset current paragraph
#                    current_cell += 1
#                    current_paragraph = 0
#                    # if we reached the end of the table, we now select the next
#                    # table
#                    if current_cell == current_table_nb_cells:
#                        current_table += 1
#                        if current_table == len(tables):
#                            pass
#                        else:
#                            current_cell = 0
#                            current_table_width = int(tables[current_table]\
#                                                      .find('tblW').get('w:w'))
#                            current_table_nb_cells = len(tables[current_table]\
#                                                         .find_all('tc'))
#                            current_nb_paragraph_in_cell = [len(x.find_all('p'))\
#                                                            for x in tables[current_table].find_all('tc')]
#                            current_cell_index = [0,0]
#            
#            text = paragraph.text
#    
#            properties = paragraph.find('w:pPr')
#    
#            style = self._try_or_none_properties(properties, 'w:pStyle', 'w:val')
#            list_lvl = self._try_or_none_properties(properties, 'w:ilvl', 'w:val')
#            list_type = self._try_or_none_properties(properties, 'w:numId', 'w:val')
#            position = self._try_or_none_properties(properties, 'jc', 'w:val')
#            bookmark = self._try_or_none_properties(paragraph, 'bookmarkStart', 'w:id')
#            
#            # We analyze the properties of each run : if runs in the same 
#            # paragraph have different style, it appears as a list in the pandas
#            runs_prop = []
#            run_prop = [ 'vertAlign', 'b', 'i', 'color', 'highlight', 'u', 'bCs', 'iCs']
#            try:
#                for run in paragraph.find_all('r'):
#                    prop = run.find('rPr')
#                    runs_output = [self._try_or_none_properties(prop, x, 'w:val') for x in run_prop]
#                    runs_prop.append(runs_output)
#                runs_prop = [x[0] if len(set(x)) == 1 else x for x in np.array(runs_prop).T]
#            except AttributeError:
#                runs_prop = [None] * 8
#    
##            tabs_prop = []
##            try: 
##                tabs = properties.find('w:tabs')
##                for i in tabs.find_all('w:tab'):
##                    tabs_prop.append([i.get('w:pos'), i.get('w:val')])
##            except AttributeError:
##                pass
#            
#            if paragraph.find('cNvPr'):
#                img_id = "_".join(['image', paragraph.find('cNvPr').get('id')])
#            else:
#                img_id = None
#    
#            characteristics = ['text', 'paragraph_style', 'list_lvl','list_type' ,\
#                               'horizontal_alignment', 'vertical_alignment',\
#                               'bold', 'italic', 'text_color', 'highlight', 'underline', \
#                               'bookmark', 'table', 'cell', 'image']
#    
#            try:
#                summary.append([text, style, list_lvl, list_type, position] +            
#                                runs_prop[:-2] + [bookmark, index_table, 
#                                cell_index, img_id])
#                
#            except IndexError: # because of issue that doesn't allow to create a list of None in a loop
#                summary.append([text, style, list_lvl, list_type, position, \
#                                None, None, None, None, None, None, bookmark, \
#                                index_table, cell_index, img_id])
#        
#        output = pd.DataFrame(summary, columns = characteristics)
#        
#        output[output == '0'] = None
#        output[output == 'none'] = None
#        
#        output.text_color[output.text_color == 'auto'] = None
#        
#        
#        self.summary = output
            
        
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
    def  _find_tables_composition(self):
        
        self.tables = self.doc.find_all('tbl')
        
        self.p_per_tables = [len(x.find_all('p')) for x in self.tables]
        self.cells_per_tables = [len(x.find_all('tc')) for x in self.tables]
        
        index_subtables = []
        self.index_parentTable_close = []
        self.index_parentTable_far = []
        max_sublevels = []
        for i in range(1,len(self.tables)):
            current_table_level = []
            for j in range(i):
                if str(self.tables[i]) in str(self.tables[j]):
                    index_subtables.append(i)
                    current_table_level.append(j)
            if current_table_level:
                self.index_parentTable_close.append(max(current_table_level))
                self.index_parentTable_far.append(min(current_table_level))
                max_sublevels.append(len(current_table_level))
        
        self.index_subtables = list(set(index_subtables))
        
        tables_sublevels = max(max_sublevels)
        
        self.tables_colnames = []
        for i in range(tables_sublevels + 1):
            self.tables_colnames.append("_".join(['table_level', str(i)]))
            
#        for i in sorted(range(len(self.index_parentTable_close)),key = lambda x: self.index_parentTable_close[x]):
#            self.p_per_tables[self.index_parentTable_close[i]] -= self.p_per_tables[self.index_subtables[i]]
#            self.cells_per_tables[self.index_parentTable_close[i]] -= self.cells_per_tables[self.index_subtables[i]]
        
        self.p_per_tables_used = self.p_per_tables.copy()
        self.cells_per_tables_used = self.cells_per_tables.copy()
    
#        tables = [[table for i, table in enumerate(tables) if i not in index_drop]]
        
        

filename = 'TESTDOC.docx'

doc = docx_reader(filename)

txt = doc._read_docx()

#doc.index_parentTable_close
#doc.index_parentTable_far

#for i in sorted(range(len(doc.index_parentTable_close)),key = lambda x: doc.index_parentTable_close[x]):
#    doc.p_per_tables[doc.index_parentTable_close[i]] -= doc.p_per_tables[doc.index_subtables[i]]
    

#doc.index_parentTable_close
#doc.index_subtables