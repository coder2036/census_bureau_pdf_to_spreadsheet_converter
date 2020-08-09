# -*- coding: utf-8 -*-
"""
Created on Fri Jun 26 16:50:11 2020

@author: gupta018
"""

from openpyxl_writing import write_sheet
from read_pdf import extract_text,read_company_id_and_name
from tkinter import filedialog
import tkinter as tk
import os
import traceback
#import pkg_resources.py2_warn # I needed this import to get pyinstaller exe to run
import threading
import re
      

# pulled from stack overflow
class Drag_and_Drop_Listbox(tk.Listbox):
  """ A tk listbox with drag'n'drop reordering of entries. """
  def __init__(self, master, **kw):
    kw['selectmode'] = tk.EXTENDED
    kw['activestyle'] = 'none'
    tk.Listbox.__init__(self, master, kw,width=50)
    self.bind('<Button-1>', self.getState, add='+')
    self.bind('<Button-1>', self.setCurrent, add='+')
    self.bind('<B1-Motion>', self.shiftSelection)
    self.curIndex = None
    self.curState = None
  def setCurrent(self, event):
    ''' gets the current index of the clicked item in the listbox '''
    self.curIndex = self.nearest(event.y)
  def getState(self, event):
    ''' checks if the clicked item in listbox is selected '''
    i = self.nearest(event.y)
    self.curState = self.selection_includes(i)
  def shiftSelection(self, event):
    ''' shifts item up or down in listbox '''
    i = self.nearest(event.y)
    if self.curState == 1:
      self.selection_set(self.curIndex)
    else:
      self.selection_clear(self.curIndex)
    if i < self.curIndex:
      # Moves up
      x = self.get(i)
      selected = self.selection_includes(i)
      self.delete(i)
      self.insert(i+1, x)
      if selected:
        self.selection_set(i+1)
      self.curIndex = i
    elif i > self.curIndex:
      # Moves down
      x = self.get(i)
      selected = self.selection_includes(i)
      self.delete(i)
      self.insert(i-1, x)
      if selected:
        self.selection_set(i-1)
      self.curIndex = i
            
      
def add_filename():
    filenames = filedialog.askopenfilenames(initialdir='',title='Select Another Input File',
                                         filetypes = [("pdf files","*.pdf")])
    for filename in filenames:
        input_filenames.insert(tk.END,filename)    

def remove_filename():
    delete_selected_from_listbox(input_filenames)

def remove_all_filenames():
    input_filenames.delete(0,tk.END)

def delete_selected_from_listbox(listbox):
    indices = listbox.curselection()
    for i in indices[::-1]:
        listbox.delete(i)

def add_folder():
    folder = tk.filedialog.askdirectory(initialdir='',title='Select Another Output Folder')
    if folder:
        output_folders.insert(tk.END,folder)  
        if only_one_folder_bool.get():
            only_one_folder()

    
def remove_folder():
    delete_selected_from_listbox(output_folders)
    
def remove_all_folders():
    output_folders.delete(0,tk.END)
    
def remove_rows():
    indices = input_filenames.curselection()
    for i in indices[::-1]:
        input_filenames.delete(i)
        output_folders.delete(i)
        outcomes.delete(i)
    indices = output_folders.curselection();
    for i in indices[::-1]:
        input_filenames.delete(i)
        output_folders.delete(i)
        outcomes.delete(i)
    indices = outcomes.curselection()
    for i in indices[::-1]:
        input_filenames.delete(i)
        output_folders.delete(i)
        outcomes.delete(i)

def create_output_filename(desired_format,company_id):
    output_filename = re.sub('{id}',str(company_id),desired_format)
    output_filename = re.sub('{year}',settings['Year'],output_filename)
    return output_filename

def convert():
    files = input_filenames.get(0,tk.END)
    folders = output_folders.get(0,tk.END)
    if len(files) == 0:
        tk.messagebox.showerror("No PDF Files Selected","Please select at least one PDF file to convert.")
        return
    if not write_to_same_folder_as_pdf_bool.get() and not only_one_folder_bool.get():
        if len(files) != len(folders):
            tk.messagebox.showerror("File-Folder Mismatch","Must select same number of input files and output folders")
            return    
    for i in range(len(files)):
        file = files[i]
        if write_to_same_folder_as_pdf_bool.get():
            folder = os.path.dirname(file)
        elif only_one_folder_bool.get():
            folder = folders[0]
        else:
            folder = folders[i]
        update_listbox(outcomes,i,'converting ' + file + ' and placing in ' + folder)
        outcome = convert_pdf(file,folder)
        update_listbox(outcomes,i,outcome)
        # convert
        # insert outcome

def write_to_same_folder_as_pdf():
    if write_to_same_folder_as_pdf_bool.get():
        output_folders.grid_remove()
        only_one_folder_bool.set(False)
        only_one_folder()
    else:
        output_folders.grid()
    
def only_one_folder():
    if only_one_folder_bool.get():
        write_to_same_folder_as_pdf_bool.set(False)
        write_to_same_folder_as_pdf()
        print(tk.END)
        while(output_folders.get(1)):
            output_folders.delete(0)
        
def update_listbox(listbox,index,text):
    listbox.delete(index)
    listbox.insert(index,text)
    root.update()
        
def clear_all():
    input_filenames.delete(0,tk.END)
    output_folders.delete(0,tk.END)
    outcomes.delete(0,tk.END)
    
def convert_pdf(pdf_path,output_folder):
    try:
        (company_id,company_name) = read_company_id_and_name(pdf_path)
        output_filename = create_output_filename(output_filename_format.get(), company_id)
        output_file_full_path = output_folder + '/'  + output_filename + '.xlsx'
        if overwrite_warning_bool.get() and os.path.isfile(output_file_full_path):
            overwrite_decision = tk.messagebox.askquestion('Overwrite File?','Are you sure you want to overwrite ' + output_file_full_path + '?' ,icon='warning')
            if overwrite_decision != 'yes':
                return 'Avoided Overwriting'
        df = extract_text(pdf_path,settings['Survey Information Spreadsheet'])  
        write_sheet(df,output_file_full_path,company_name)
    except PermissionError:
        print('permission error')
        traceback.print_exc()
        return 'Close ' + output_file_full_path + ' before trying again'
    except Exception as e:
        print('error')
        traceback.print_exc()
        return 'Failed'
    print('normal')
    return 'Completed: ' + output_file_full_path
    
def read_settings_to_dict():
    settings_file = open('converter_settings.txt','r')
    content = settings_file.read()
    lines = content.split('\n')
    settings = {}
    for line in lines:
        if line.startswith('#') or not line or not line.strip():
            continue
        key = line.split('=')[0].strip()
        value = line.split('=')[1].strip()
        settings[key] = value
    return settings

settings = read_settings_to_dict()   
    
    
root = tk.Tk()
root.grid_columnconfigure(0,weight=1)

row = 0

frame = tk.Frame(root)
frame.grid(row=row,column=0,sticky='ew')
frame.grid_columnconfigure(0,weight=1,uniform='third')
frame.grid_columnconfigure(1,weight=1,uniform='third')
frame.grid_columnconfigure(2,weight=1,uniform='third')

pdf_files_label = tk.Label(frame,text='PDF Files',borderwidth=1,relief='solid')
folders_label = tk.Label(frame,text='Output Folders',borderwidth=1,relief='solid')
outcomes_label = tk.Label(frame,text='Outcomes',borderwidth=1,relief='solid')
pdf_files_label.grid(column=0,row=row,sticky='nsew')
folders_label.grid(column=1,row=row,sticky='nsew')
outcomes_label.grid(column=2,row=row,sticky='nsew')

row += 1

input_filenames = tk.Listbox(frame,selectmode='extended',exportselection=0) # may want 'multiple' selectmode
output_folders = Drag_and_Drop_Listbox(frame,selectmode='extended',exportselection=0)
outcomes = tk.Listbox(frame,selectmode='extended',exportselection=0)
input_filenames.grid(column=0,row=row,sticky='nsew')
output_folders.grid(column=1,row=row,sticky='nsew')
outcomes.grid(column=2,row=row,sticky='nsew')
    
# =============================================================================
# for i in range(100):
#     input_filenames.insert(tk.END,'filasdfsfasdfasdfasdfasdfasdfasdfasdfasdfefilasdfsfasdfasdfasdfasdfasdfasdfasdfasdfefilasdfsfasdfasdfasdfasdfasdfasdfasdfasdfefilasdfsfasdfasdfasdfasdfasdfasdfasdfasdfe' + str(i))
#     output_folders.insert(tk.END,'folder' + str(i))
#     outcomes.insert(tk.END,'outcomes' + str(i))
# =============================================================================
    

def yview(*args):
    input_filenames.yview(*args)
    output_folders.yview(*args)
    outcomes.yview(*args)
    
scrollbar_y = tk.Scrollbar(frame,command=yview)
scrollbar_y.grid(row=row,column=4,sticky='ns')
input_filenames.config(yscrollcommand=scrollbar_y.set)
output_folders.config(yscrollcommand=scrollbar_y.set)
outcomes.config(yscrollcommand=scrollbar_y.set)
#not sure if these are necessary^^

row += 1

def set_up_x_scroll(frame,element,column):
    scrollbar_x = tk.Scrollbar(frame,orient='horizontal',command=element.xview)
    element.config(xscrollcommand=scrollbar_x.set)
    scrollbar_x.grid(column=column,row=row,sticky='ew')
    
set_up_x_scroll(frame,input_filenames,0)
set_up_x_scroll(frame,output_folders,1)
set_up_x_scroll(frame,outcomes,2)


row +=1
add_filenames_button = tk.Button(frame,text='Add PDF File(s)',command = add_filename)
add_filenames_button.grid(column=0,row=row,pady=(20,0))
add_folders_button = tk.Button(frame,text='Add Output Folder',command=add_folder)
add_folders_button.grid(column=1,row=row,pady=(20,0))

row += 1

remove_filenames_button = tk.Button(frame,text='Remove Selected PDF Files',command = remove_filename)
remove_filenames_button.grid(column=0,row=row)
remove_folders_button = tk.Button(frame,text='Remove Selected Output Folders',command=remove_folder)
remove_folders_button.grid(column=1,row=row)

row+=1
remove_all_filenames_button = tk.Button(frame,text="Remove All PDF Files",command=remove_all_filenames)
remove_all_filenames_button.grid(column=0,row=row)
remove_all_folders_button = tk.Button(frame,text="Remove All Output Folders",command=remove_all_folders)
remove_all_folders_button.grid(column=1,row=row)

row += 1
overwrite_warning_bool = tk.BooleanVar()
overwrite_warning_bool.set(True)
overwrite_toggle = tk.Checkbutton(frame,text='Ask before overwriting a file',var=overwrite_warning_bool)
overwrite_toggle.grid(column=0,row=row)
write_to_same_folder_as_pdf_bool = tk.BooleanVar()
write_to_same_folder_as_pdf_toggle = tk.Checkbutton(frame,text='Write Spreadsheet to same folder as pdf file',var=write_to_same_folder_as_pdf_bool,command=write_to_same_folder_as_pdf)
write_to_same_folder_as_pdf_bool.set(False)
write_to_same_folder_as_pdf_toggle.grid(column=1,row=row)

row += 1
only_one_folder_bool = tk.BooleanVar()
only_one_folder_bool.set(False)
only_one_folder_toggle = tk.Checkbutton(frame,text='Write All Outputs to One Folder',var=only_one_folder_bool,command=only_one_folder)
only_one_folder_toggle.grid(column=1,row=row)

row += 1
remove_rows_button = tk.Button(frame,text='Remove Selected Rows',command=remove_rows)
remove_rows_button.grid(column=1,row=row,pady=(30,0))

row += 1
clear_button = tk.Button(frame,text='Remove All Rows',command=clear_all)
clear_button.grid(column=1,row=row)

print(frame.winfo_width())
row+=1
output_filename_format_lbl = tk.Label(frame,text="Provide output file name format below.\n ( {id} uses the company's ID and {year} pulls from converter_settings.txt )")
output_filename_format_lbl.grid(column=1,row=row,pady=(20,0))

row+=1
output_filename_format = tk.StringVar()
output_filename_format_entry = tk.Entry(frame,textvar=output_filename_format,justify='center')
output_filename_format.set("{id}_CompanyCalendar_{year}")
output_filename_format_entry.grid(column=1,row=row,sticky='ew')

row += 1
convert_button = tk.Button(frame,text='Convert PDFs to Spreadsheets',command=convert)
convert_button.grid(column=1,row=row,pady=20)

root.attributes('-topmost',True)
root.attributes('-topmost',False)
root.focus_force()

root.mainloop()
