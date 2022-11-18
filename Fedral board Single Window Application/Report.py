import tkinter as tk
from tkinter import ttk
# from docx2pdf import convert
from PIL import Image,ImageTk

wind = tk.Tk()  # wind window name
wind.title('windboard By Danish')  # title Name
# wind = Tk()
# wind.geometry("1600x1600")
# bg = PhotoImage(file = "Web 1920 – 1(1).png")

# # Show image using label
# label1 = Label( wind, image = bg)
# label1.place(x = 0, y = 0)
# wind.iconbitmap('add icon link And Directory name')    # icon add

# function coding start 


exp = " "          # global variable 
# showing all data in display 

def press(num):
    global exp
    exp=exp + str(num)
    equation.set(exp)
# end 


# function clear button

def clear():
    global exp
    exp = " "
    equation.set(exp)

# end 


# Enter Button Work Next line Function

def action():
  exp = " Next Line : "
  equation.set(exp)

# end function coding









# Size window size
wind.geometry('1010x250')         # normal size
wind.maxsize(width=1010, height=250)      # maximum size
wind.minsize(width= 1010 , height = 250)     # minimum size
# end window size

wind.configure(bg = 'green')    #  add background color

# entry box
equation = tk.StringVar()
Dis_entry = ttk.Entry(wind,state= 'readonly',textvariable = equation)
Dis_entry.grid(rowspan= 1 , columnspan = 100, ipadx = 999 , ipady = 20)
# end entry box

# add all button line wise 

# First Line Button

q = ttk.Button(wind,text = '1' , width = 6, command = lambda : press('1'))
q.grid(row = 1 , column = 0, ipadx = 6 , ipady = 10)

w = ttk.Button(wind,text = '2' , width = 6, command = lambda : press('2'))
w.grid(row = 1 , column = 1, ipadx = 6 , ipady = 10)

E = ttk.Button(wind,text = '3' , width = 6, command = lambda : press('3'))
E.grid(row = 1 , column = 2, ipadx = 6 , ipady = 10)


A = ttk.Button(wind,text = '4' , width = 6, command = lambda : press('4'))
A.grid(row = 2 , column = 0, ipadx = 6 , ipady = 10)

clear = ttk.Button(wind,text = 'Clear' , width = 6, command = clear)
clear.grid(row = 1 , column = 5, ipadx = 6 , ipady = 10)



S = ttk.Button(wind,text = '5' , width = 6, command = lambda : press('5'))
S.grid(row = 2 , column = 1, ipadx = 6 , ipady = 10)

D = ttk.Button(wind,text = '6' , width = 6, command = lambda : press('6'))
D.grid(row = 2 , column = 2, ipadx = 6 , ipady = 10)


Z = ttk.Button(wind,text = '7' , width = 6, command = lambda : press('7'))
Z.grid(row = 3 , column = 0, ipadx = 6 , ipady = 10)


X = ttk.Button(wind,text = '8' , width = 6, command = lambda : press('8'))
X.grid(row = 3 , column = 1, ipadx = 6 , ipady = 10)


C = ttk.Button(wind,text = '9' , width = 6, command = lambda : press('9'))
C.grid(row = 3 , column = 2, ipadx = 6 , ipady = 10)






wind.mainloop()  # using ending point