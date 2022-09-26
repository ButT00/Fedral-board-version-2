import codecs
import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
import win32api
import docx2txt
import aspose.words as aw


root = tk.Tk();
root.geometry("550x750") 
#Title
root.title("Fedral Board")

my_button = Button(text="Print", relief="raised", font=("Times New Roman", 22),command=lambda : [Print() ] )
my_button.place(relx = 0.6,
                    rely = 0.7,
                    anchor = 'center',
                    width=100
                    )



def Print():
        
        # Ask for file (Which you want to print)
        file_to_print = "RESULT CANCELLATION CERTIFICATE.docx"
        doc = aw.Document("RESULT CANCELLATION CERTIFICATE.docx")

        # Save as PDF
        doc.save("PDF.pdf")
        
        
        
        if file_to_print:
            
            # Print Hard Copy of File
            win32api.ShellExecute(0,              # NULL since it's not associated with a window
                "print",        # execute the "print" verb defined for the file type
                file_to_print,  # path to the document file to print
                None,           #no parameters, since the target is a document file
                ".",            #current directory, same as NULL here
                0) 
            



filename = 'RESULT CANCELLATION CERTIFICATE.docx'



import aspose.words as aw

doc = aw.Document('RESULT CANCELLATION CERTIFICATE.docx')
doc.save("Output.txt")
            

configfile = Text(root, wrap=WORD, width=175, height=100)
configfile.pack(fill="none", expand=TRUE)

filename = 'Output.txt'
with open(filename,"rt", encoding="latin-1") as f:
    configfile.insert(INSERT, f.read())          


my_button = Button(text="Verified", relief="raised", font=("Times New Roman", 22),command=lambda : [Print() ] )
my_button.place(relx = 0.3,
                    rely = 0.8,
                    anchor = 'center',
                    width=100
                    )

my_button = Button(text="Cancel", relief="raised", font=("Times New Roman", 22),command=lambda : [root.destroy() ] )
my_button.place(relx = 0.5,
                    rely =0.8,
                    anchor = 'center',
                    width=100
                    )
    
  

            
root.mainloop()         
            
