from tkinter import *
from ConfirmDocument import *
import tkinter as tk
from PIL import Image, ImageTk
import os
from datetime import datetime
import shutil
import pypyodbc as pyodbc
import pandas as pd
import os
from openpyxl import load_workbook
from docx import Document
from docxtpl import DocxTemplate


#  connection with DB ssc
try:
    cnxn_str = ("Driver={SQL Server Native Client 11.0};"
            "Server=DBSERVER;"
            "Database=ssc;"
            "UID=kiosk01;"
            "PWD=123;")

    cnxn = pyodbc.connect(cnxn_str)


 
except:
    
    def OpenScanDocumentFile():
        os.system("python ScanDocument.py")
    window = Tk()
    window.geometry('1600x1600')
    window.title("DB connection Failed")
    Error = tk.Label(window, text="Error", font=("Microsoft JhengHei", 18))
    Error.place(relx = 0.4,
                    rely = 0.2,
                    anchor = 'center') 
    Database_error = tk.Label(window, text="Database Connection Failed Please Connect LAN Cable", font=("Microsoft JhengHei", 16))
    Database_error.place(relx = 0.4,
                    rely = 0.3,
                    anchor = 'center') 
    my_button_retry = Button(text="Retry", relief="raised", font=("Times New Roman", 16),  command=lambda : [ OpenScanDocumentFile(), ] )
    my_button_retry.place(relx = 0.4,
                    rely = 0.4,
                    anchor = 'center',
                    )  
    window.mainloop()
    
# cnxn_str = ("Driver={SQL Server Native Client 11.0};"
#             "Server=DBSERVER;"
#             "Database=ssc;"
#             "UID=kiosk01;"
#             "PWD=123;")

# cnxn = pyodbc.connect(cnxn_str)



# connection with db ssc ledger 
try:
    cnxn_str_2 = ("Driver={SQL Server Native Client 11.0};"
                "Server=DBSERVER;"
                "Database=ssc_LEDGER;"
                "UID=kiosk01;"
                "PWD=123;")

    cnxn_2 = pyodbc.connect(cnxn_str_2)
except:    
    def OpenScanDocumentFile():
        os.system("python ScanDocument.py")
    window = Tk()
    window.geometry('1600x1600')
    window.title("DB Connection Failed")
    Error = tk.Label(window, text="Error", font=("Microsoft JhengHei", 18))
    Error.place(relx = 0.4,
                    rely = 0.2,
                    anchor = 'center') 
    my_button_retry = Button(text="Retry", relief="raised", font=("Times New Roman", 16),  command=lambda : [window.destroy(), OpenScanDocumentFile()] )
    my_button_retry.place(relx = 0.4,
                    rely = 0.4,
                    anchor = 'center',
                    )
    Database_error = tk.Label(window, text="Database Connection Failed Please Connect LAN Cable", font=("Microsoft JhengHei", 16))
    Database_error.place(relx = 0.4,
                    rely = 0.3,
                    anchor = 'center')  
    window.mainloop()
    

cursor_2 = cnxn_2.cursor()

cursor = cnxn.cursor()

# Current date.
today = datetime.today()
Dated = today.date()


nbr = '9010110004'
               
Name_display = ''' SELECT name FROM ZReg WHERE reg_no =?'''

cursor.execute(Name_display, [nbr])
    
Name_display_O = cursor.fetchone()[0]




# Condition to select the Certificate

# select reg_no from challan where ch_no = '981273123'

# certificate = select certificate from challan where ch_no = '89371293'

# if(certifiacte == "MIGRATION CERTIFICATE")

# For Migration Certificate        

M_C_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

Name_M_C = M_C_Name.fetchone()[0]

M_C_DOB = cursor.execute("SELECT dob from ZReg WHERE reg_no = '9010110004'")

DOB_M_C = M_C_DOB.fetchone()[0]

M_C_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

fname_M_C = M_C_fname.fetchone()[0]

M_C_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

reg_no_M_C = M_C_reg_no.fetchone()[0]

M_C_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

roll_no_M_C = M_C_roll_no.fetchone()[0]

M_C_Year = cursor.execute("Select year from ZReg WHERE reg_no = '9010110004'")

Year_M_C = M_C_Year.fetchone()[0]

M_C_Institution = cursor.execute("select inst_desc from ZReg where reg_no = '9010110004'")

Institution_M_C = M_C_Institution.fetchone()[0]


# /////// Completed Migration Certificate

# # for Result Cancelation Certificate

R_C_C_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

Name_R_C_C = R_C_C_Name.fetchone()[0]

R_C_C_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

fname_R_C_C = R_C_C_fname.fetchone()[0]

R_C_C_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

roll_no_R_C_C = R_C_C_roll_no.fetchone()[0]

R_C_C_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

reg_no_R_C_C = R_C_C_reg_no.fetchone()[0]

# Completed Result Cancelation Certificate

# For Application form for migration request

AMR_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

Name_AMR = AMR_Name.fetchone()[0]

AMR_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

fname_AMR = AMR_fname.fetchone()[0]

AMR_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

roll_no_AMR = AMR_roll_no.fetchone()[0]

AMR_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

reg_no_AMR = AMR_reg_no.fetchone()[0]

AMR_Year = cursor.execute("Select year from ZReg WHERE reg_no = '9010110004'")

Year_AMR = AMR_Year.fetchone()[0]










# completed Application form for migration certificate

# For making Entry only in integer
class Lotfi(tk.Entry):
    def __init__(self, master=None, **kwargs):
        self.var = tk.StringVar()
        tk.Entry.__init__(self, master, textvariable=self.var, **kwargs)
        self.old_value = '' 
        self.var.trace('w', self.check)
        self.get, self.set = self.var.get, self.var.set

    def check(self, *args):
        if self.get().isdigit(): 
            # the current value is only digits; allow this
            self.old_value = self.get() 
        else:
            # there's non-digit characters in the input; reject this 
            self.set(self.old_value)

def disable_event():
       pass
   


# Starting of our software
def start(root):
    #Title
    root.title("Fedral Board")
    
    root.protocol("WM_DELETE_WINDOW", disable_event)
   
    
    #dimensions
    canvas = tk.Canvas(root, width=1300, height=1100)
    canvas.grid(columnspan=3, rowspan=3)



    #logo 
    logo = Image.open('logo.jpg')
    logo = ImageTk.PhotoImage(logo)
    logo_label = tk.Label(image=logo)
    logo_label.image = logo
    logo_label.grid(column=1, row=0)




    #instructions
    instructions = tk.Label(root, text="Please scan barcode From your Chalan form",bg='#121c76'
    , relief="raised",font=("Times New Roman", 22))
    instructions.place(relx = 0.5,
                    rely = 0.3,
                    anchor = 'center')
    instructions2 = tk.Label(root, text="Or",bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    instructions2.place(relx = 0.5,
                    rely = 0.4,
                    anchor = 'center')

    instructions3 = tk.Label(root, text="Type chalan number",bg='#add8e6',relief="raised", font=("Times New Roman", 22))
    instructions3.place(relx = 0.5,
                    rely = 0.5,
                    anchor = 'center')

    From_entry = Lotfi(root, width=25)
    
    

    From_entry.place(relx = 0.5,
                    rely = 0.6,
                    anchor = 'center',
                    width=200,
                    height=30
                    )
  
    
   
    # For closing the program
    def Closing_Window():
        
        win = Tk()
        win.geometry('600x300')
        win.title("Closing Window")
        instructions_closing = tk.Label(win, text="NOT FOR STUDENTS", font=("Microsoft JhengHei", 14))
        instructions_closing.place(relx = 0.3,
                        rely = 0.3,
                        anchor = 'center')
        
        pass_var=Entry(win, show="*")
        pass_var.place(relx = 0.3,
                        rely = 0.4,
                        anchor = 'center',
                        width=200,
                        height=30
                        )
        
        def close_root():
            pass_value = pass_var.get()
            if(pass_value == "close"):
                root.destroy()
                win.destroy()
            else:
                pass
        
        #Create a button to close the window
        btnclose = tk.Button(win, text ="Click here to Close",command=close_root)
        btnclose.place(relx = 0.3,
                        rely = 0.5,
                        anchor = 'center',
                    
                        )
    
    
    
    
        
        win.mainloop()  
    btn = Button(root, text="X", command = Closing_Window,bg='red')
    btn.place(relx = 0.9,
                    rely = 0.1,
                    anchor = 'center',
                    width=50,
                    height=30
                    
                    ) 
      
        
#   function for getting the challan number from the textbox s
    def challan_Entry():
        
        challan_No = From_entry.get()
    
        nbr = challan_No
               
        command = ''' SELECT name FROM ZReg WHERE reg_no =?'''

        cursor.execute(command, [nbr])
    
        data = cursor.fetchone()
        
        print(data)
            

    # from one screen to another

    my_button = Button(text="Check" ,relief="raised",font=("Times New Roman", 22),  command=lambda : [Making_document(), challan_Entry(), change(root) ] )
    my_button.place(relx = 0.5,
                    rely = 0.7,
                    anchor = 'center',
                    width=300,
                    height=86
                    )
    

def Making_document():
    # Open files
    main_path = r"C:\Users\hp\Desktop\Fedral board project\Fedral Board\Fedral board Project\Fedral board Single Window Application"
    
    #Condition for selecting the template file which print 
    # if(certificate == "MIGRATION CERTIFICATE"):
        
    # template_path = os.path.join(main_path, 'MIGRATION CERTIFIACTE_templ.docx')
    # workbook_path = os.path.join(main_path, 'Template_data.xlsx')

    # workbook = load_workbook(workbook_path)
    # template = DocxTemplate(template_path)
    # worksheet = workbook["Input"]

    # to_fill_in = {'Candidate_name' : None,
    #             'Dated' : None,
    #             'DOB': None,
    #             'Father_name': None,
    #             'Registration_no' : None,
    #             'Roll_no' : None,
    #             'Examination': None,
    #             'Session': None,
    #             'year': None,
    #             'Status': None,
    #             'Institution': None
                
    #             }

   
   

    # to_fill_in['Candidate_name'] = Name_M_C
    # to_fill_in['DOB'] = DOB_M_C
    # to_fill_in['Dated'] = Dated
    # to_fill_in['Father_name'] = fname_M_C
    # to_fill_in['Registration_no'] = reg_no_M_C
    # to_fill_in['Roll_no'] = roll_no_M_C
    # to_fill_in['Examination'] = "Fedral board"
    # to_fill_in['Session'] = "Final Session"
    # to_fill_in['year'] =  Year_M_C
    # to_fill_in['Status'] = "pass"
    # to_fill_in['Institution'] = Institution_M_C
        
        
    # # Fill in all the keys defined in the word document using the dictionary.
    # # The keys in de word document are identified by the {{}}symbols.
    # template.render(to_fill_in)
    # # Output the file to a docx document.
    # filename = 'MIGRATION CERTIFIACTE.docx'
    # filled_path = os.path.join(main_path, filename)
    # template.save(filled_path)
    # print("Done with MIGRATION CERTIFIACTE.docx")
    
    # completed Migration certificate Filling 
    
    # if(certificate == "RESULT CANCELLATION CERTIFICATE"):
    
    template_path = os.path.join(main_path, 'RESULT CANCELLATION CERTIFICATE_templ.docx')
    workbook_path = os.path.join(main_path, 'Template_data.xlsx')

    # workbook = load_workbook(workbook_path)
    template = DocxTemplate(template_path)
    # worksheet = workbook["Input"]

    to_fill_in = {'Candidate_name' : None,
                  'Father_name': None,
                  'Roll_no' : None,
                  'Registration_no' : None,
                  
                }

    to_fill_in['Candidate_name'] = Name_R_C_C
    to_fill_in['Dated'] = Dated
    to_fill_in['Father_name'] = fname_R_C_C
    to_fill_in['Registration_no'] = reg_no_R_C_C
    to_fill_in['Roll_no'] = roll_no_R_C_C
    
        
        
    # Fill in all the keys defined in the word document using the dictionary.
    # The keys in de word document are identified by the {{}}symbols.
    template.render(to_fill_in)
    # Output the file to a docx document.
    filename = 'RESULT CANCELLATION CERTIFICATE.docx'
    filled_path = os.path.join(main_path, filename)
    template.save(filled_path)
    print("Done with RESULT CANCELLATION CERTIFICATE.docx")
    
    
    
    
    
    
    

def change(root):
    root.destroy()
    fun()
    
def call():
    root = Tk();
    start(root)
    root.mainloop()
    
# This function is calling the window
if __name__ == '__main__':
    call() 
    

    




