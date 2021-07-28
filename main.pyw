#!venv/scripts/pythonw.exe
from tkinter import *
from tkinter.filedialog import asksaveasfilename
from tkinter import ttk
from docx import Document
from dictionaries import antigen, pcr

lists = [
    'Report Antigen',
    'Report Swab PCR'
]

def print_report():
    mask = [
        ("Word files","*.doc *.docx"),  
        ("All files","*.*")]  
    filename = asksaveasfilename(filetypes=mask, defaultextension = '.docx') 
    
    #REDIRECT SELECTED COMBOBOX TO TEMPLATES HERE
    if template.get()=='Report Antigen':
        antigen.getdata(filename,labno.get())
    if template.get()=='Report Swab PCR':
        pcr.getdata(filename,labno.get())
    
    #END REDIRECT TO TEMPLATES

#init main wind
root = Tk()
root.title("Print Result Report")
root.geometry("450x200")
root.resizable(0,0)

#init components
labnoLabel = Label(root,text="Lab No.",anchor="e",font=("Courier",11))
#pidLabel = Label(root,text="PID",anchor="e",font=("Courier",11))
#patientnameLabel = Label(root,text="Patient Name",anchor="e",font=("Courier",11))
templateLabel = Label(root,text="Template",anchor="e",font=("Courier",11))
labno = Entry(root, width=15, borderwidth=2)
#pid = Entry(root, width=15, borderwidth=2)
#patientname = Entry(root, width=15, borderwidth=2)
template = ttk.Combobox(root,values=lists,width=30)
template.current(0)
printButton = Button(root,text="Print",width=7,height=1,border=2,font=("Courier",11,"bold"))
infoLabel = Label(root,anchor="e",font=("Courier",11))

#positioning components
labnoLabel.grid(row=1,column=1,padx=2,pady=5,sticky=W+E)
labno.grid(row=1,column=2,padx=5,pady=3,sticky=W+E)
#pidLabel.grid(row=2,column=1,padx=2,pady=5,sticky=W+E)
#pid.grid(row=2,column=2,padx=2,pady=5,sticky=W+E)
#patientnameLabel.grid(row=3,column=1,padx=2,pady=5,sticky=W+E)
#patientname.grid(row=3,column=2,padx=5,pady=3,sticky=W+E)
templateLabel.grid(row=4,column=1,padx=2,pady=5,sticky=W+E)
template.grid(row=4,column=2,padx=5,pady=3,sticky=W+E)
printButton.grid(row=5,column=2,padx=4,pady=5,sticky=W+E)
infoLabel.grid(row=6,column=1,columnspan=2,padx=2,pady=5,sticky=W+E)

root.grid_rowconfigure(0, minsize=30)
root.grid_columnconfigure(0, minsize=50)

printButton.config(command=print_report)

root.mainloop()
