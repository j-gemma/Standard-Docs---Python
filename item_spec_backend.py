from openpyxl import load_workbook
from openpyxl import cell
from docx import Document
import re
from datetime import date
from tkinter import messagebox
from docxcompose.composer import Composer

class AqItem:
    
    def __init__(self, item_no, mfr, model, category, spec, accs):
        self.item_no = item_no
        self.mfr = mfr
        self.model = model
        self.category = category
        self.spec = spec
        self.accs = accs
        self.line1 = "ITEM {} - {}\n".format(self.item_no.upper(), self.category.upper())
        self.line2 = ''
      
class AqAcc:

    def __init__(self, mfr, model_no, spec):
        self.mfr = mfr
        self.model_no = model_no
        self.spec = spec   
        

def main(xl_file, save_as, save_path, to_compose):

    print('working')
    ws, column_dict = get_columns(xl_file)
    items = make_spec_dict(ws, column_dict)
    items, error_item_nos = clean_data(items)
    item_spec = make_spec(items)
    item_spec.save('temp.docx')

    if to_compose == 'nocompose':
        pass
    else:
        item_spec = compose()
    
    item_spec.save(save_path + "/" + save_as + '.docx')
    return error_item_nos

def copy_header_footer(header_doc, item_spec):
    # Open the source and target documents
    source = Document(header_doc)
    target = item_spec

    source_header = source.sections[0].header
    target_header = item_spec.sections[0].header

    for paragraph in source_header.paragraphs:
        print(paragraph.text + 'Style: ' + paragraph.style.name)
        new_para = target_header.add_paragraph(paragraph.text, paragraph.style)
        new_para_format = new_para.paragraph_format
        copy_format = paragraph.paragraph_format

        new_para_format.alignment = copy_format.alignment
        new_para_format.first_line_indent = copy_format.first_line_indent
        new_para_format.keep_together = copy_format.keep_together
        new_para_format.keep_with_next = copy_format.keep_with_next
        new_para_format.left_indent = copy_format.left_indent
        new_para_format.line_spacing = copy_format.line_spacing
        new_para_format.line_spacing_rule = copy_format.line_spacing_rule
        new_para_format.page_break_before = copy_format.page_break_before
        new_para_format.right_indent = copy_format.right_indent
        new_para_format.space_after = copy_format.space_after
        new_para_format.space_before = copy_format.space_before
        for tab_stop in copy_format.tab_stops:
            new_para_format.tab_stops.add_tab_stop(tab_stop.position, tab_stop.alignment)
        new_para_format.widow_control = copy_format.widow_control
    return target

def make_spec(items):
    
    template_path = r'Q:\Standard Documents\Templates' \
                    + '\\' \
    + "SECTION 114000 FOODSERVICE ITEM SPEC.docx" 
    template = Document(template_path)

    for key, item in items.items():
        if item.category == "OPEN NUMBER":
            para = template.add_paragraph(item.line1)
            para.style= template.styles['AqItem']
        else:
            para = template.add_paragraph(item.line1 + item.line2)
            para.style= template.styles['AqItem']

            para = template.add_paragraph(item.spec)
            para.style = template.styles['ItemSpec']

        if len(item.accs) > 0:  
            for acc in item.accs:
                para = template.add_paragraph(acc.spec)
                para.style = template.styles['Accessory']

    for paragraph in template.paragraphs:
        for run in paragraph.runs:
            if " Gc" in run.text:
                # Replace the text and preserve formatting
                new_text = run.text.replace(" Gc", " GC")
                run.text = new_text

    return template

def clean_data(items):

    print('item.item_no')
    

    punct_patt = r"NIKEC(,/)"
    by_who = r"NIKEC[/,]\s?((?:(?:BY|by|By)|(?:IN|in|In))\s((?:[^/\s]+(?:/[^/\s]+)*)+))"
    nikec_string = 'This item is not in the kitchen equipment contract.'
    error_item_nos = []
    
    for key, item in items.items():
        print(item.item_no)

        if item.mfr == 'maxi':
            item.mfr = 'Maxi Movers'

        if item.category != 'OPEN NUMBER':
            item.line2 = '{} Model {}\n'.format(item.mfr, item.model)
        if item.spec is not None:
            if item.spec[:5] == "NIKEC":
                punct = item.spec[5] if len(item.spec) >= 6 else ' '
                print("'{}'".format(punct))
                if punct == ',':
                    print('comma')
                    punct += ' '
                
                match_obj = re.search(by_who, item.spec)  # Search once and store the match object
                if match_obj:
                    nikec_by = match_obj.group(1)
                else: 
                    nikec_by = 'FIX ME IN AUTOQUOTES'
                    error_item_nos.append(item.item_no)

                item.line2 = 'NIKEC' + punct + nikec_by.title() + '\n'
                item.spec = nikec_string
                item.accs = []
                              
        elif item.model is not None:
            if item.model[:5] == "NIKEC":
              if item.spec[:5] == "NIKEC":
                  pass
              else:
                punct = item.model[5]
                if punct == ',':
                    punct += ' '
                    
                match_obj = re.search(by_who, item.model)  # Search once and store the match object
                if match_obj:
                    nikec_by = match_obj.group(1)
                else: 
                    nikec_by = 'FIX ME IN AUTOQUOTES'
                    error_item_nos.append(item.item_no)

                item.line2 = 'NIKEC' + punct + nikec_by.title() + '\n'
                item.spec = nikec_string
                item.accs = []
            
        if len(item.accs) >= 1:
            for acc in item.accs:
                if acc.mfr == item.mfr:
                    if acc.model_no is not None:
                        acc.spec = "Model {} {}".format(acc.model_no, acc.spec)
                    else:
                        pass
                    
                elif acc.mfr is not None:
                    if acc.model_no is not None:
                        acc.spec = "{} Model {} {}".format(acc.mfr, \
                                                           acc.model_no,\
                                                           acc.spec)
                    else:
                        pass              
    return items, error_item_nos

def compose():

    front = r"Q:\Standard Documents\Templates\FOODSERVICE GENERAL SPEC FRONT.docx"
    back = r"Q:\Standard Documents\Templates\FOODSERVICE GENERAL SPEC BACK.docx"
    middle = r"temp.docx"

    composer = Composer(Document(front))
    composer.append(Document(middle))
    composer.append(Document(back))
    
    return composer.doc

def get_columns(filename):

    columns = {}

    find_columns = ['ItemNo','Mfr', 'Model', 'Category', 'Spec']
    wb = load_workbook(filename = filename)
    ws = wb['Spreadsheet']
    header_row = ws[1]
    for i in range(len(find_columns)):
        for c in header_row:
            if c.value == find_columns[i]:
                columns[find_columns[i]] = c.column_letter

    return ws, columns

def make_spec_dict(ws, column_dict):
    item_no_col_index = column_dict['ItemNo']
    mfr_col_index = column_dict['Mfr']
    model_col_index = column_dict['Model']
    category_col_index = column_dict['Category']
    spec_col_index = column_dict['Spec']
    item_no_col = ws[item_no_col_index]
    spec_dict = {}
    
    for row_num in range(1, len(item_no_col)):
        row = str(row_num + 1)
        item_cell = item_no_col[row_num]
        item_no = item_cell.value
        mfr = ws[mfr_col_index + row].value
        model = ws[model_col_index + row].value
        category = ws[category_col_index + row].value
        spec = ws[spec_col_index + row].value
        accs = []
        
        if item_cell.value != None and (item_cell.fill.fgColor.rgb == 'FFFFFFFF'
                                    or item_cell.fill.fgColor.rgb == '00FFFFFF'):

            tempidx = row_num + 1

            while (tempidx < len(item_no_col)
            and item_no_col[tempidx].value == None):
                if(item_no_col[tempidx].fill.fgColor.rgb == 'FFFAFAD2' or \
                   item_no_col[tempidx].fill.fgColor.rgb == '00FAFAD2'):
                    row = str(tempidx + 1)
                    accs.append(AqAcc(ws[mfr_col_index + row].value,
                                      ws[model_col_index + row].value,
                                      ws[spec_col_index + row].value))
                    tempidx += 1
                else:
                    tempidx += 1
                row = str(tempidx + 1)
            
            

            spec_dict[item_no] = AqItem(item_no, mfr, model,
                                        category, spec, accs)
    return spec_dict
