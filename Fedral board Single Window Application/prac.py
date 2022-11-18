  # try:
    #     challan_No = from_entry.get()

    #     nbr = challan_No
                
    #     command = ''' select verified from income where bank_chalan_number = ? '''


    #     cursor_5.execute(command, [nbr])
    #     pay_or_not = cursor_5.fetchone()[0]
    #     print(pay_or_not)
    #     if(pay_or_not== "1"):
    #         root = Tk()
    #         root.title("ALready Taken")
    #         root.geometry("600x400")

    #         instructions1_ = Label(root, text="You Already ", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #         instructions1_.place(relx = 0.4,
    #                             rely = 0.2,
    #                             anchor = 'center')

    #         instructions2_ = Label(root, text="Taken", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #         instructions2_.place(relx = 0.4,
    #                             rely = 0.3,
    #                             anchor = 'center')

    #         instructions3_ = Label(root, text="This servise", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #         instructions3_.place(relx = 0.4,
    #                             rely = 0.4,
    #                             anchor = 'center')

    #         root.mainloop()
# except:
    # pay or not check open 
    # try:
    #     challan_No = from_entry.get()

    #     nbr = challan_No
                
    #     command = ''' select bank_chalan_number from income where bank_chalan_number= ? '''


    #     cursor_5.execute(command, [nbr])
    #     pay_or_not = cursor_5.fetchone()[0]
    #     print(pay_or_not)
    # except:
    #     pay_or_not = 0
    # challan_No = from_entry.get()

    # nbr = challan_No
            
    # command_taken = ''' select verified from income where bank_chalan_number = ?  '''


    # cursor_5.execute(command_taken, [nbr])
    # taken = cursor_5.fetchone()[0]
    # print(taken)
    # if(taken == '1'):
    #     root = Tk()
    #     root.title("ALready Taken")
    #     root.geometry("600x400")

    #     instructions1_ = Label(root, text="You Already ", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #     instructions1_.place(relx = 0.4,
    #                         rely = 0.2,
    #                         anchor = 'center')

    #     instructions2_ = Label(root, text="Taken", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #     instructions2_.place(relx = 0.4,
    #                         rely = 0.3,
    #                         anchor = 'center')

    #     instructions3_ = Label(root, text="This servise", bg='#add8e6', relief="raised", font=("Times New Roman", 22))
    #     instructions3_.place(relx = 0.4,
    #                         rely = 0.4,
    #                         anchor = 'center')

    #     root.mainloop()
    # if(pay_or_not == challan_No):
        # pay or not check close