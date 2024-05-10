#import python-docx
from docx import Document
import re
from tkinter import filedialog, messagebox

prefix_info = "[INFO] "
prefix_error = "[ERROR] "

# Create small GUI to select the docx file
filename = filedialog.askopenfilename()

#if filename is not docx file, exit the program
if not filename.endswith('.docx'):
    messagebox.showerror('Error', 'Please select a docx file')
    print(prefix_error + 'Please select a docx file')
    exit()

#open the docx file
doc = Document(filename)

year = None

#get first year in the docx file and print it
for para in doc.paragraphs:
    if re.search(r'\b\d{4}\b', para.text):
        year = re.search(r'\b\d{4}\b', para.text).group()
        break

#find also in tables and replace with the next year
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if re.search(r'\b\d{4}\b', cell.text):
                for match in re.finditer(r'\b\d{4}\b', cell.text):
                    if match.group() == year:
                        cell.text = cell.text.replace(year, str(int(year)+1))


#find all years like the found year+1 and replace them with the next year
for para in doc.paragraphs:
    if re.search(r'\b\d{4}\b', para.text):
        for match in re.finditer(r'\b\d{4}\b', para.text):
            if match.group() == str(int(year)+1):
                para.text = para.text.replace(str(int(year)+1), str(int(year)+2))


#find all years like the found year and replace them with the next year
for para in doc.paragraphs:
    if re.search(r'\b\d{4}\b', para.text):
        for match in re.finditer(r'\b\d{4}\b', para.text):
            if match.group() == year:
                # higlight the change in the docx filr
                para.text = para.text.replace(year, str(int(year)+1))




#find all years like the found year-1 and replace them with the next year
for para in doc.paragraphs:
    if re.search(r'\b\d{4}\b', para.text):
        for match in re.finditer(r'\b\d{4}\b', para.text):
            if match.group() == str(int(year)-1):
                para.text = para.text.replace(str(int(year)-1), year)

#save the docx file with the updated years
doc.save('updated_'+filename.split('/')[-1])

#dialog box to show the file is updated and saved and where it is saved
messagebox.showinfo('Info', 'File updated and saved as updated_'+filename.split('/')[-1])

print(prefix_info + 'File updated and saved as updated_'+filename.split('/')[-1])

#dont close the program
input('Press Enter to close the program')


