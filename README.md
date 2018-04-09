# docx-reader

Simple and easy analysis of information contained in a .docx file paragraph by paragraph.

Any suggestions is welcome regarding the format of the script.


# 09.04.2018

At the moment, it returns a pandas dataframe as a summary of the document, with columns defining characteristics of each row (paragraph).

Please note :
  - Cells in table may contain multiple paragraphs
  - A Paragraph may contain multiple run : multiple style inside the same paragraph
      If that is the case, the corresponding cell in the pandas DataFrame output is a list
  - There is no installer yet



