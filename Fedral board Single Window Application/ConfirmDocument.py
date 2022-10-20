import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
import win32api
from tkinter import filedialog
from pyautogui import alert
import time
import os
from tkPDFViewer import tkPDFViewer as pdf 
import shutil
from ScanDocument import *

 

def OpenScanDocumentFile():
        os.system("python ScanDocument.py")

def disable_event():
       pass


def fun():
    root = tk.Tk();
    #Title
    root.title("Fedral Board")
    root.protocol("WM_DELETE_WINDOW", disable_event)

    #dimensions
    canvas = tk.Canvas(root, width=800, height=600)
    canvas.grid(columnspan=3, rowspan=3)



    #logo 
    logo = Image.open('logo.jpg')
    logo = ImageTk.PhotoImage(logo)
    logo_label = tk.Label(image=logo)
    logo_label.image = logo
    logo_label.grid(column=1, row=0)
    
    # creating object of ShowPdf from tkPDFViewer. 
    v1 = pdf.ShowPdf() 


  
# # Adding pdf location and width and height. 
#     v2 = v1.pdf_view(
#                  pdf_location= "PDF.pdf") 
 
#     v2.place(relx = 0.2,
#                     rely = 0.3,
#                     anchor = 'center')
    
    # # Intruction to verify our document
    # instructions_V = tk.Label(root, text="* PLease Verify your information", bg='#add8e6', relief="raised", font=("Times New Roman", 28))
    # instructions_V.place(relx = 0.2,
    #                 rely = 0.3,
    #                 anchor = 'center')
    
    # instructions_V_2 = tk.Label(root, text="1.Is your name is : "+ Name_display_O , bg='#add8e6', relief="raised", font=("Times New Roman", 18))
    # instructions_V_2.place(relx = 0.2,
    #                 rely = 0.4,
    #                 anchor = 'center')
    
    # instructions_V_3 = tk.Label(root, text="2.Is your Reg is : " , bg='#add8e6', relief="raised", font=("Times New Roman", 18))
    # instructions_V_3.place(relx = 0.2,
    #                 rely = 0.5,
    #                 anchor = 'center')
    
    # instructions_V_3_ = tk.Label(root, text="3. Certificate which you want to print is : " , bg='#add8e6', relief="raised", font=("Times New Roman", 18))
    # instructions_V_3_.place(relx = 0.2,
    #                 rely = 0.6,
    #                 anchor = 'center')
    
    # instructions_V_4 = tk.Label(root, text="If all the information statisfies" ,bg='#add8e6', relief="raised", font=("Times New Roman", 20))
    # instructions_V_4.place(relx = 0.2,
    #                 rely = 0.7,
    #                 anchor = 'center')
    
    # instructions_V_5 = tk.Label(root, text="Then print otherwise exit" , bg='#add8e6', relief="raised", font=("Times New Roman", 20))
    # instructions_V_5.place(relx = 0.2,
    #                 rely = 0.8,
    #                 anchor = 'center')
    
    
     



    #instructions
    instructions = tk.Label(root, text="Do you want ", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    instructions.place(relx = 0.5,
                    rely = 0.3,
                    anchor = 'center')

    instructions2 = tk.Label(root, text="to Print this",bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    instructions2.place(relx = 0.5,
                    rely = 0.4,
                    anchor = 'center')

    instructions2 = tk.Label(root, text="Document", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    instructions2.place(relx = 0.5,
                    rely = 0.5,
                    anchor = 'center')
    
    
    # Create a textbox
    # my_text = Text(root, height=480, width=480)
    # my_text = tk.Label(root,  bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    # my_text.place(relx = 0.3,
    #                 rely = 0.1,
    #                 anchor = 'center')
    

    # my_text.delete(1.0, END)

    # # Open our pdf file

    # # Grab the filename of the pdf file
    # open_file = "PDF.pdf"

    # # Check to see if there is a file
        
    # # Open the pdf file
    # pdf_file = PyPDF2.PdfFileReader(open_file)
    # # Set the page to read
    # page = pdf_file.getPage(0)
    # # Extract the text from the pdf file
    # page_stuff = page.extractText()
    # # Add text to textbox
    # my_text.insert(1.0, page_stuff)
    
    #Check button
    def Print():
        
        
        
        
        if os.path.exists("MIGRATION CERTIFIACTE.docx"):
            
          
            
            # Ask for file (Which you want to print)
            file_to_print = "MIGRATION CERTIFIACTE.docx"
            
            
            
            if file_to_print:
                
                # Print Hard Copy of File
                win32api.ShellExecute(0,              # NULL since it's not associated with a window
                    "print",        # execute the "print" verb defined for the file type
                    file_to_print,  # path to the document file to print
                    None,           #no parameters, since the target is a document file
                    ".",            #current directory, same as NULL here
                    0)              # SW_HIDE passed to app associated with the file type 
                
        if os.path.exists("RESULT CANCELLATION CERTIFICATE.docx"):
            file_to_print = "RESULT CANCELLATION CERTIFICATE.docx"
            
            
            
            if file_to_print:
                
                # Print Hard Copy of File
                win32api.ShellExecute(0,              # NULL since it's not associated with a window
                    "print",        # execute the "print" verb defined for the file type
                    file_to_print,  # path to the document file to print
                    None,           #no parameters, since the target is a document file
                    ".",            #current directory, same as NULL here
                    0)
            
            
        
# this function is used when we print the document 

        time.sleep(10)
        if os.path.exists("RESULT CANCELLATION CERTIFICATE.docx"):
            os.remove("RESULT CANCELLATION CERTIFICATE.docx") # one file at a time
        if os.path.exists("RESULT CANCELLATION CERTIFICATE.pdf"):
            os.remove("RESULT CANCELLATION CERTIFICATE.pdf") # one file at a time
        if os.path.exists("RESULT CANCELLATION CERTIFICATE.png"):
            os.remove("RESULT CANCELLATION CERTIFICATE.png") # one file at a time
        if os.path.exists("MIGRATION CERTIFIACTE.docx"):
            os.remove("MIGRATION CERTIFIACTE.docx") # one file at a time
        if os.path.exists("MIGRATION CERTIFIACTE.png"):
            os.remove("MIGRATION CERTIFIACTE.png") # one file at a time
        if os.path.exists("MIGRATION CERTIFIACTE.pdf"):
            os.remove("MIGRATION CERTIFIACTE.pdf") # one file at a time
            
        
        
        
        
                          
        
        
        
        
        
    def NotPrint():
        print("not print")
        
    # def OpenScanDocumentFile():
    #     os.system("python ScanDocument.py")
        

        
    my_button = Button(text="Print", relief="raised", font=("Times New Roman", 22),command=lambda : [Print(),root.destroy(), OpenScanDocumentFile() ] )
    my_button.place(relx = 0.4,
                    rely = 0.7,
                    anchor = 'center',
                    width=100
                    )

    my_button2 = Button(text="Exit", relief="raised", font=("Times New Roman", 22),  command=lambda : [root.destroy(), OpenScanDocumentFile()] )
    my_button2.place(relx = 0.6,
                    rely = 0.7,
                    anchor = 'center',
                    width=100
                    )
    
    

    
    
    root.mainloop()




    

    