from docx import Document 
from datetime import date
from tkinter import messagebox

def main(areas, proj_num, proj_title, directory, location, cover_date):
    TEMPLATE_PATH = r"Q:\Standard Documents\Templates - Do not edit or move these files!\Cutbook Covers\CUTBOOK COVER.docx"
    print("template: {} \nproject: {} \nnumber:{} \ndirectory: {} \nlocation: {}"\
          .format(TEMPLATE_PATH, proj_title, proj_num, directory, location))
    for k,v in areas.items():
        print("\nkey: {} \nvalue: {}".format(k, v))
    print('making covers') 

    for key, value in areas.items():
        replacements = init_replacements()
        
        replacements['<<Area>>'] = value.upper()
        replacements['<<Abbr>>'] = key.upper()
        print('key {}'.format(key))
        replacements['<<Title>>'] = proj_title.upper()
        replacements['<<Location>>'] = location.upper()
        replacements['<<Date>>'] = cover_date.upper()
        replacements['<<Num>>'] = proj_num
        print('dict: {}'.format(replacements))
        
        template = Document(TEMPLATE_PATH)
        filename = '{}\\{}.docx'.format(directory, value.upper())
        
        for k, v in replacements.items():
            print('key: {}, value: {}\n'.format(k, v))
            word = k
            new_word = str(v)
            template = find_and_replace(template, word, new_word)

        # Save the modified document to a new file
        template.save(filename)
    
    print('done')


def init_replacements():
   replacements = {
        '<<Area>>' : '',
        '<<Abbr>>' : '',
        '<<Title>>' : '',
        '<<Location>>' : '',
        '<<Num>>' : '',
        '<<Date>>' : '',
        }
   return replacements




def find_and_replace(doc, old_text, new_text):
  """
  Finds and replaces all occurrences of `old_text` with `new_text` in the document.

  Args:
    doc: A docx document object.
    old_text: The string to be replaced.
    new_text: The string to replace `old_text` with.
  """

  # Iterate through paragraphs
  i= 0
  for paragraph in doc.paragraphs:
    
    for run in paragraph.runs:
      if len(run.text) > 0:
          print("run number {}: {}".format(i, run.text))
          run.text = run.text.replace(old_text, new_text)
          i += 1

  i=0

  return doc
