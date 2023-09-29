from tkinter import *
from tkinter import ttk
from tkinter import filedialog

from docx import Document
from docx2python import docx2python
from docx2python.iterators import iter_paragraphs
 
root = Tk()
root.title("exTRACKTOR")
root.geometry("250x200")
 
root.grid_rowconfigure(index=0, weight=1)
root.grid_columnconfigure(index=0, weight=1)
root.grid_columnconfigure(index=1, weight=1)
 
text_editor = Text()
text_editor.grid(column=0, columnspan=2, row=0)

def open_file(event=None):
    filename = filedialog.askopenfilename()

    docx_content = docx2python(filename)
    footnotes = '\n'.join(iter_paragraphs(docx_content.footnotes))
    splited = footnotes.splitlines()
    filtered = list(filter(None, splited))

    for i in range(len(filtered)):
        filtered[i] = filtered[i].replace('footnote', '')
        filtered[i] = filtered[i].replace('\t', '')
        filtered[i] = filtered[i].replace('\t', '')
        filtered[i] = filtered[i][3:]
        filtered[i] = filtered[i].lstrip()

    docx_content.close()
    document = Document()
    for item in filtered:
        document.add_paragraph(item)
    
    document.save(f"{filename} (Only Footnotes).docx")
    
    root.destroy()
 
open_button = ttk.Button(text="Открыть файл", command=open_file)
open_button.grid(column=0, row=1, sticky=NSEW, padx=10)
 
root.mainloop()
