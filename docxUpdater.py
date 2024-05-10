# Imports
from docx import Document
import re
from tkinter import filedialog, messagebox

# Prefixes for the print statements
prefix_info = "[INFO] "
prefix_error = "[ERROR] "

# Dialog box to select the docx file
filename = filedialog.askopenfilename()

# If the filename is empty, show an error message and exit the program or if the file is not a docx file
if not filename.endswith('.docx'):
    messagebox.showerror('Error', 'Please select a docx file')
    print(prefix_error + 'Please select a docx file')
    input('Press Enter to close the program')
    exit()

# Create a Document object from the docx file
doc = Document(filename)

# Find the first year in the docx file and store it in the variable year
year = None
for para in doc.paragraphs:
    if re.search(r'\b\d{4}\b', para.text):
        year = re.search(r'\b\d{4}\b', para.text).group()
        break

# Find the year in the table and replace it with the next year
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            if re.search(r'\b\d{4}\b', cell.text):
                for match in re.finditer(r'\b\d{4}\b', cell.text):
                    if match.group() == year:
                        cell.text = cell.text.replace(year, str(int(year) + 1))

# Find the year in the paragraphs and replace it with the next year
for i in range(1, -2, -1):
    for para in doc.paragraphs:
        if re.search(r'\b\d{4}\b', para.text):
            for match in re.finditer(r'\b\d{4}\b', para.text):
                if match.group() == str(int(year) + i):
                    para.text = para.text.replace(str(int(year) + i), str(int(year) + i + 1))

# Find the year in the filename and replace it with the next year
filename = filename.replace(year, str(int(year) + 1))

# Save the updated docx file
doc.save(filename.split('/')[-1])

# Dialog box to show the file is updated and saved
messagebox.showinfo('Info', 'File updated and saved as ' + filename.split('/')[-1])

# Print the file is updated and saved and the path of the updated file (absolute path)
print(prefix_info + 'File updated and saved as ' + filename.split('/')[-1])
print(prefix_info + 'Path: ' + filename + '\n')

# Wait for enter key to close the program
input('Press Enter to close the program')
