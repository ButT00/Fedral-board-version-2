#Import the required Libraries
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
from docx2pdf import convert
from pdf2image import convert_from_path
from PIL import Image,ImageTk
from pdf2image import convert_from_path
from docx2pdf import convert
from PIL import Image,ImageTk

check = 1
# connecting with the payments database for getting the reg_no on the refference of challan no
import mysql.connector
from mysql.connector import Error

try:
    connection = mysql.connector.connect(host='192.168.100.2',
                                         database='fee_challans',
                                         user='yaseen2',
                                         password='fVum*.NODLS]w_6F')
    if connection.is_connected():
        db_Info = connection.get_server_info()
        print("Connected to MySQL Server version ", db_Info)
        cursor_P = connection.cursor()
        # cursor_P.execute("SELECT reg_no FROM `payments` WHERE challan_no = 771666175600 ")
        # record = cursor_P.fetchone()[0]
        # print(record)

except Error as e:
    print("Error while connecting to MySQL", e)

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
    
    

try:
    cnxn_str_3 = ("Driver={SQL Server Native Client 11.0};"
            "Server=DBSERVER;"
            "Database=hssc;"
            "UID=kiosk01;"
            "PWD=123;")

    cnxn_3 = pyodbc.connect(cnxn_str_3)


 
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
    
    
try:
    cnxn_str_4 = ("Driver={SQL Server Native Client 11.0};"
            "Server=DBSERVER;"
            "Database=Hssc_LEDGER;"
            "UID=kiosk01;"
            "PWD=123;")

    cnxn_4 = pyodbc.connect(cnxn_str_4)


 
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


cursor_2 = cnxn_2.cursor()

cursor_3 = cnxn_3.cursor()

cursor_4 = cnxn_4.cursor()

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

# # For Migration Certificate        

# M_C_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

# Name_M_C = M_C_Name.fetchone()[0]

# M_C_DOB = cursor.execute("SELECT dob from ZReg WHERE reg_no = '9010110004'")

# DOB_M_C = M_C_DOB.fetchone()[0]

# M_C_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

# fname_M_C = M_C_fname.fetchone()[0]

# M_C_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

# reg_no_M_C = M_C_reg_no.fetchone()[0]

# M_C_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

# roll_no_M_C = M_C_roll_no.fetchone()[0]

# M_C_Year = cursor.execute("Select year from ZReg WHERE reg_no = '9010110004'")

# Year_M_C = M_C_Year.fetchone()[0]

# M_C_Institution = cursor.execute("select inst_desc from ZReg where reg_no = '9010110004'")

# Institution_M_C = M_C_Institution.fetchone()[0]


# /////// Completed Migration Certificate

# # for Result Cancelation Certificate

# R_C_C_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

# Name_R_C_C = R_C_C_Name.fetchone()[0]

# R_C_C_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

# fname_R_C_C = R_C_C_fname.fetchone()[0]

# R_C_C_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

# roll_no_R_C_C = R_C_C_roll_no.fetchone()[0]

# R_C_C_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

# reg_no_R_C_C = R_C_C_reg_no.fetchone()[0]

# Completed Result Cancelation Certificate

# For Application form for migration request

# AMR_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

# Name_AMR = AMR_Name.fetchone()[0]

# AMR_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

# fname_AMR = AMR_fname.fetchone()[0]

# AMR_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

# roll_no_AMR = AMR_roll_no.fetchone()[0]

# AMR_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

# reg_no_AMR = AMR_reg_no.fetchone()[0]

# AMR_Year = cursor.execute("Select year from ZReg WHERE reg_no = '9010110004'")

# Year_AMR = AMR_Year.fetchone()[0]

# cursor_P.execute("SELECT challan_no FROM payments WHERE head_code = 0300313")
# a = cursor_P.fetchone()[0]
# print(a)





# gloabal variables used in this system



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
    instructions = tk.Label(root, text="Please scan barcode From your Chalan form",bg='#add8e6'
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
    # def challan_Entry():
        
        
    #     challan_No = From_entry.get()
    
    #     nbr = challan_No
               
    #     command = ''' SELECT reg_no FROM payments WHERE challan_no =%s'''

    #     cursor_P.execute(command, [nbr])
    #     global reg_no_
        
    #     reg_no_ = cursor_P.fetchone()[0]
        
    #     # # for Result Cancelation Certificate

    #     R_C_C_Name = cursor.execute("SELECT name from ZReg WHERE reg_no = '9010110004'")

    #     Name_R_C_C = R_C_C_Name.fetchone()[0]

    #     R_C_C_fname = cursor.execute("Select fname from ZReg WHERE reg_no = '9010110004'")

    #     fname_R_C_C = R_C_C_fname.fetchone()[0]

    #     R_C_C_roll_no = cursor_2.execute("select roll_no from ZLedgerII where reg_no = '9010110004'")

    #     roll_no_R_C_C = R_C_C_roll_no.fetchone()[0]

    #     R_C_C_reg_no = cursor.execute("Select reg_no from ZReg where fname = 'KARIM KHAN MARWAT'")

    #     reg_no_R_C_C = R_C_C_reg_no.fetchone()[0]

# Completed Result Cancelation Certificate
    
        
        
        
        # command2 = ''' SELECT reg_no FROM payments WHERE challan_no =%s'''
    
        
        
        
        # print(reg_no_)
        

        
        # print(data)
            

    # from one screen to another

    my_button = Button(text="Check" ,relief="raised",font=("Times New Roman", 22),  command=lambda : [  Making_document() ] )
    my_button.place(relx = 0.5,
                    rely = 0.7,
                    anchor = 'center',
                    width=300,
                    height=86
                    )
    
    
    def Making_document():
        
        
        challan_No = From_entry.get()
    
        nbr = challan_No
               
        command = ''' SELECT reg_no FROM payments WHERE challan_no =%s'''
        

        cursor_P.execute(command, [nbr])
    
        
        reg_no_ = cursor_P.fetchone()[0]
        print(reg_no_)
        
        
        certificate_name = ''' SELECT head_code FROM `payments` WHERE challan_no  =%s'''
        cursor_P.execute(certificate_name, [nbr])
        certificate_name_code =  cursor_P.fetchone()[0]
        print(certificate_name_code)
        
        
        payment_status = ''' SELECT payment_status FROM `payments` WHERE challan_no =%s'''
        cursor_P.execute(payment_status, [nbr])
        status_P =  cursor_P.fetchone()[0]
        
        
        
        
        
        
        
        
         # Open files
        main_path = r"C:\Users\hp\Desktop\Fedral board project\Fedral Board\Fedral board Project\Fedral board Single Window Application"
        
         #Condition for selecting the template file which print 
        
        if certificate_name_code == '02000250':
          
        
            # # For Application form for migration request

            # AMR_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_Name, [reg_no_])
            

            # Name_AMR = cursor.fetchone()[0]

            # AMR_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_fname, [reg_no_])

            # fname_AMR = cursor.fetchone()[0]

            # AMR_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            # cursor_2.execute(AMR_roll_no, [reg_no_])

            # roll_no_AMR = cursor_2.fetchone()[0]

            # AMR_reg_no = '''Select reg_no from ZReg where fname = ?'''
            # cursor.execute(AMR_reg_no, [fname_AMR])

            # reg_no_AMR = cursor.fetchone()[0]

            # AMR_Year = '''Select year from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_Year, [reg_no_])

            # Year_AMR = cursor.fetchone()[0]
            
            # # For Migration Certificate        

            M_C_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            cursor.execute(M_C_Name, [reg_no_])

            Name_M_C = cursor.fetchone()[0]
        
            M_C_DOB = '''SELECT dob from ZReg WHERE reg_no = ?'''
            cursor.execute(M_C_DOB, [reg_no_])

            DOB_M_C = cursor.fetchone()[0]

            M_C_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            cursor.execute(M_C_fname, [reg_no_])

            fname_M_C = cursor.fetchone()[0]

            M_C_reg_no = '''Select reg_no from ZReg where fname = ?'''
            cursor.execute(M_C_reg_no, [fname_M_C])

            reg_no_M_C = cursor.fetchone()[0]

            M_C_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            cursor_2.execute(M_C_roll_no, [reg_no_])

            roll_no_M_C = cursor_2.fetchone()[0]

            M_C_Year = '''Select year from ZReg WHERE reg_no = ?'''
            cursor.execute(M_C_Year, [reg_no_])

            Year_M_C = cursor.fetchone()[0]

            M_C_Institution = '''select inst_desc from ZReg where reg_no = ?'''
            cursor.execute(M_C_Institution, [reg_no_])

            Institution_M_C = cursor.fetchone()[0]
            
            template_path = os.path.join(main_path, 'MIGRATION CERTIFIACTE_templ.docx')
            workbook_path = os.path.join(main_path, 'Template_data.xlsx')

            workbook = load_workbook(workbook_path)
            template = DocxTemplate(template_path)
            worksheet = workbook["Input"]

            to_fill_in = {'Candidate_name' : None,
                        'Dated' : None,
                        'DOB': None,
                        'Father_name': None,
                        'Registration_no' : None,
                        'Roll_no' : None,
                        'Examination': None,
                        'Session': None,
                        'year': None,
                        'Status': None,
                        'Institution': None
                        
                        }

        
        

            to_fill_in['Candidate_name'] = Name_M_C
            to_fill_in['DOB'] = DOB_M_C
            to_fill_in['Dated'] = Dated
            to_fill_in['Father_name'] = fname_M_C
            to_fill_in['Registration_no'] = reg_no_M_C
            to_fill_in['Roll_no'] = roll_no_M_C
            to_fill_in['Examination'] = "Fedral board"
            to_fill_in['Session'] = "Final Session"
            to_fill_in['year'] =  Year_M_C
            to_fill_in['Status'] = "pass"
            to_fill_in['Institution'] = Institution_M_C
                
                
            # Fill in all the keys defined in the word document using the dictionary.
            # The keys in de word document are identified by the {{}}symbols.
            template.render(to_fill_in)
            # Output the file to a docx document.
            filename = 'MIGRATION CERTIFIACTE.docx'
            filled_path = os.path.join(main_path, filename)
            template.save(filled_path)
            print("Done with MIGRATION CERTIFIACTE.docx")
            
            convert("MIGRATION CERTIFIACTE.docx", "MIGRATION CERTIFIACTE.pdf")
            
            images = convert_from_path("MIGRATION CERTIFIACTE.pdf", 500,poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
            for i, image in enumerate(images):
                fname = 'MIGRATION CERTIFIACTE'+'.png'
                image.save(fname, "PNG")
                
            # Python3 program to convert docx to 
                

            #Import the required Libraries

            #Create an instance of tkinter frame
            wind = tk.Toplevel()

            #Set the geometry of tkinter frame
            wind.geometry("1600x1600")

            #Create a canvas
            canvas= Canvas(wind, width= 900, height= 900)
            canvas.pack()

            #Load an image in the script
            img= (Image.open("MIGRATION CERTIFIACTE.png"))

            #Resize the Image using resize method
            resized_image= img.resize((900,705), Image.Resampling.LANCZOS)
            new_image= ImageTk.PhotoImage(resized_image)

            #Add image to the Canvas Items
            canvas.create_image(10,10, anchor=NW, image=new_image)

            instructions1_ = Label(wind, text="Is This That document", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions1_.place(relx = 0.1,
                                rely = 0.2,
                                anchor = 'center')


            instructions2_ = Label(wind, text="you want", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions2_.place(relx = 0.1,
                                rely = 0.3,
                                anchor = 'center')


            instructions3_ = Label(wind, text="to print", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions3_.place(relx = 0.1,
                                rely = 0.4,
                                anchor = 'center')


            btn_ = Button(wind, text="Yes", command = lambda:[wind.destroy,change(root)])
            btn_.place(relx = 0.1,
                                rely = 0.5,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 

            btn2_ = Button(wind, text="No", command = wind.destroy)
            btn2_.place(relx = 0.1,
                                rely = 0.6,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 
            
            wind.mainloop()
        
        # completed Migration certificate Filling 
        
        if certificate_name_code == '03000350':
              
        
            # # For Application form for migration request

            # AMR_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_Name, [reg_no_])
            

            # Name_AMR = cursor.fetchone()[0]

            # AMR_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_fname, [reg_no_])

            # fname_AMR = cursor.fetchone()[0]

            # AMR_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            # cursor_2.execute(AMR_roll_no, [reg_no_])

            # roll_no_AMR = cursor_2.fetchone()[0]

            # AMR_reg_no = '''Select reg_no from ZReg where fname = ?'''
            # cursor.execute(AMR_reg_no, [fname_AMR])

            # reg_no_AMR = cursor.fetchone()[0]

            # AMR_Year = '''Select year from ZReg WHERE reg_no = ?'''
            # cursor.execute(AMR_Year, [reg_no_])

            # Year_AMR = cursor.fetchone()[0]
            
            # # For Migration Certificate        

            M_C_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            cursor_3.execute(M_C_Name, [reg_no_])

            Name_M_C = cursor_3.fetchone()[0]
        
            # M_C_DOB = '''SELECT dob from ZReg WHERE reg_no = ?'''
            # cursor_3.execute(M_C_DOB, [reg_no_])

            # DOB_M_C = cursor_3.fetchone()[0]

            M_C_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            cursor_3.execute(M_C_fname, [reg_no_])

            fname_M_C = cursor_3.fetchone()[0]

            M_C_reg_no = '''Select reg_no from ZReg where fname = ?'''
            cursor_3.execute(M_C_reg_no, [fname_M_C])

            reg_no_M_C = cursor_3.fetchone()[0]

            M_C_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            cursor_4.execute(M_C_roll_no, [reg_no_])

            roll_no_M_C = cursor_4.fetchone()[0]

            M_C_Year = '''Select year from ZReg WHERE reg_no = ?'''
            cursor_3.execute(M_C_Year, [reg_no_])

            Year_M_C = cursor_3.fetchone()[0]

            M_C_Institution = '''select inst_desc from ZReg where reg_no = ?'''
            cursor_3.execute(M_C_Institution, [reg_no_])

            Institution_M_C = cursor_3.fetchone()[0]
            
            template_path = os.path.join(main_path, 'MIGRATION CERTIFIACTE_templ.docx')
            workbook_path = os.path.join(main_path, 'Template_data.xlsx')

            workbook = load_workbook(workbook_path)
            template = DocxTemplate(template_path)
            worksheet = workbook["Input"]

            to_fill_in = {'Candidate_name' : None,
                        'Dated' : None,
                        # 'DOB': None,
                        'Father_name': None,
                        'Registration_no' : None,
                        'Roll_no' : None,
                        'Examination': None,
                        'Session': None,
                        'year': None,
                        'Status': None,
                        'Institution': None
                        
                        }

        
        

            to_fill_in['Candidate_name'] = Name_M_C
            # to_fill_in['DOB'] = DOB_M_C
            to_fill_in['Dated'] = Dated
            to_fill_in['Father_name'] = fname_M_C
            to_fill_in['Registration_no'] = reg_no_M_C
            to_fill_in['Roll_no'] = roll_no_M_C
            to_fill_in['Examination'] = "Fedral board"
            to_fill_in['Session'] = "Final Session"
            to_fill_in['year'] =  Year_M_C
            to_fill_in['Status'] = "pass"
            to_fill_in['Institution'] = Institution_M_C
                
                
            # Fill in all the keys defined in the word document using the dictionary.
            # The keys in de word document are identified by the {{}}symbols.
            template.render(to_fill_in)
            # Output the file to a docx document.
            filename = 'MIGRATION CERTIFIACTE.docx'
            filled_path = os.path.join(main_path, filename)
            template.save(filled_path)
            print("Done with MIGRATION CERTIFIACTE.docx")
            
            convert("MIGRATION CERTIFIACTE.docx", "MIGRATION CERTIFIACTE.pdf")
            
            images = convert_from_path("MIGRATION CERTIFIACTE.pdf", 500,poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
            for i, image in enumerate(images):
                fname = 'MIGRATION CERTIFIACTE'+'.png'
                image.save(fname, "PNG")
                
            # Python3 program to convert docx to 
                

            #Import the required Libraries

            #Create an instance of tkinter frame
            wind = tk.Toplevel()

            #Set the geometry of tkinter frame
            wind.geometry("1600x1600")

            #Create a canvas
            canvas= Canvas(wind, width= 900, height= 900)
            canvas.pack()

            #Load an image in the script
            img= (Image.open("MIGRATION CERTIFIACTE.png"))

            #Resize the Image using resize method
            resized_image= img.resize((900,705), Image.Resampling.LANCZOS)
            new_image= ImageTk.PhotoImage(resized_image)

            #Add image to the Canvas Items
            canvas.create_image(10,10, anchor=NW, image=new_image)

            instructions1_ = Label(wind, text="Is This That document", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions1_.place(relx = 0.1,
                                rely = 0.2,
                                anchor = 'center')


            instructions2_ = Label(wind, text="you want", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions2_.place(relx = 0.1,
                                rely = 0.3,
                                anchor = 'center')


            instructions3_ = Label(wind, text="to print", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions3_.place(relx = 0.1,
                                rely = 0.4,
                                anchor = 'center')


            btn_ = Button(wind, text="Yes", command = lambda:[wind.destroy,change(root)])
            btn_.place(relx = 0.1,
                                rely = 0.5,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 

            btn2_ = Button(wind, text="No", command = wind.destroy)
            btn2_.place(relx = 0.1,
                                rely = 0.6,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 
            
            wind.mainloop()
        
        # completed Migration certificate Filling 
        
        
        
        
        
        
        
       
        # # for Result Cancelation Certificate
        
        if certificate_name_code == '02000213':
            

            R_C_C_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            cursor.execute(R_C_C_Name, [reg_no_])

            Name_R_C_C = cursor.fetchone()[0]

            R_C_C_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            cursor.execute(R_C_C_fname, [reg_no_])
            

            fname_R_C_C = cursor.fetchone()[0]

            R_C_C_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            cursor_2.execute(R_C_C_roll_no, [reg_no_])

            roll_no_R_C_C = cursor_2.fetchone()[0]

            R_C_C_reg_no = '''Select reg_no from ZReg where fname = ?'''
            cursor.execute(R_C_C_reg_no, [fname_R_C_C])

            reg_no_R_C_C = cursor.fetchone()[0]
            
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
            
    
            
            # Converting docx present in the same folder
            # as the python file
            # convert("RESULT CANCELLATION CERTIFICATE.docx")
            
            # Converting docx specifying both the input
            # and output paths
            convert("RESULT CANCELLATION CERTIFICATE.docx", "RESULT CANCELLATION CERTIFICATE.pdf")
            
            images = convert_from_path("RESULT CANCELLATION CERTIFICATE.pdf", 500,poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
            for i, image in enumerate(images):
                fname = 'RESULT CANCELLATION CERTIFICATE'+'.png'
                image.save(fname, "PNG")
                
            # Python3 program to convert docx to 
                

            #Import the required Libraries

            #Create an instance of tkinter frame
            wind = tk.Toplevel()

            #Set the geometry of tkinter frame
            wind.geometry("1600x1600")

            #Create a canvas
            canvas= Canvas(wind, width= 900, height= 900)
            canvas.pack()

            #Load an image in the script
            img= (Image.open("RESULT CANCELLATION CERTIFICATE.png"))

            #Resize the Image using resize method
            resized_image= img.resize((900,705), Image.Resampling.LANCZOS)
            new_image= ImageTk.PhotoImage(resized_image)

            #Add image to the Canvas Items
            canvas.create_image(10,10, anchor=NW, image=new_image)

            instructions1_ = Label(wind, text="Is This That document", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions1_.place(relx = 0.1,
                                rely = 0.2,
                                anchor = 'center')


            instructions2_ = Label(wind, text="you want", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions2_.place(relx = 0.1,
                                rely = 0.3,
                                anchor = 'center')


            instructions3_ = Label(wind, text="to print", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions3_.place(relx = 0.1,
                                rely = 0.4,
                                anchor = 'center')


            btn_ = Button(wind, text="Yes", command = lambda:[wind.destroy,change(root)])
            btn_.place(relx = 0.1,
                                rely = 0.5,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 

            btn2_ = Button(wind, text="No", command = wind.destroy)
            btn2_.place(relx = 0.1,
                                rely = 0.6,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 
            
            wind.mainloop()
            
            # # for Result Cancelation Certificate
        
        if certificate_name_code == '03000313':
            

            R_C_C_Name = '''SELECT name from ZReg WHERE reg_no = ?'''
            cursor_3.execute(R_C_C_Name, [reg_no_])

            Name_R_C_C = cursor_3.fetchone()[0]

            R_C_C_fname = '''Select fname from ZReg WHERE reg_no = ?'''
            cursor_3.execute(R_C_C_fname, [reg_no_])
            

            fname_R_C_C = cursor_3.fetchone()[0]

            R_C_C_roll_no = '''select roll_no from ZLedgerII where reg_no = ?'''
            cursor_4.execute(R_C_C_roll_no, [reg_no_])

            roll_no_R_C_C = cursor_4.fetchone()[0]

            R_C_C_reg_no = '''Select reg_no from ZReg where fname = ?'''
            cursor_3.execute(R_C_C_reg_no, [fname_R_C_C])

            reg_no_R_C_C = cursor_3.fetchone()[0]
            
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
            
    
            
            # Converting docx present in the same folder
            # as the python file
            # convert("RESULT CANCELLATION CERTIFICATE.docx")
            
            # Converting docx specifying both the input
            # and output paths
            convert("RESULT CANCELLATION CERTIFICATE.docx", "RESULT CANCELLATION CERTIFICATE.pdf")
            
            images = convert_from_path("RESULT CANCELLATION CERTIFICATE.pdf", 500,poppler_path=r'C:\Program Files\poppler-0.68.0\bin')
            for i, image in enumerate(images):
                fname = 'RESULT CANCELLATION CERTIFICATE'+'.png'
                image.save(fname, "PNG")
                
            # Python3 program to convert docx to 
                

            #Import the required Libraries

            #Create an instance of tkinter frame
            wind = tk.Toplevel()

            #Set the geometry of tkinter frame
            wind.geometry("1600x1600")

            #Create a canvas
            canvas= Canvas(wind, width= 900, height= 900)
            canvas.pack()

            #Load an image in the script
            img= (Image.open("RESULT CANCELLATION CERTIFICATE.png"))

            #Resize the Image using resize method
            resized_image= img.resize((900,705), Image.Resampling.LANCZOS)
            new_image= ImageTk.PhotoImage(resized_image)

            #Add image to the Canvas Items
            canvas.create_image(10,10, anchor=NW, image=new_image)

            instructions1_ = Label(wind, text="Is This That document", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions1_.place(relx = 0.1,
                                rely = 0.2,
                                anchor = 'center')


            instructions2_ = Label(wind, text="you want", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions2_.place(relx = 0.1,
                                rely = 0.3,
                                anchor = 'center')


            instructions3_ = Label(wind, text="to print", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
            instructions3_.place(relx = 0.1,
                                rely = 0.4,
                                anchor = 'center')


            btn_ = Button(wind, text="Yes", command = lambda:[wind.destroy,change(root)])
            btn_.place(relx = 0.1,
                                rely = 0.5,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 

            btn2_ = Button(wind, text="No", command = wind.destroy)
            btn2_.place(relx = 0.1,
                                rely = 0.6,
                                anchor = 'center',
                                width=50,
                                height=30
                                
                                ) 
            
            wind.mainloop()
            
            
            

        
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
    
    





