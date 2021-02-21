from tkinter import *

font = 'Tahoma'
font_size = 10
label_width = 30
field_width = 90

ged_file_name = ''
initial_family = 0
title_page_required = ''
title1 = ''
title2 = ''
title3 = ''
title4 = ''
contents_required = ''
index_required = ''
page_size = ''
document_images_required = ''
website_path = ''
country_to_remove1 = ''
country_to_remove2 = ''
person1 = 0
person2 = 0
person3 = 0
person4 = 0
spaces = '                                            '

field = ["GED File", "Page Size", "Initial Family", "Title Page Required", "Contents Required", \
        "Index Required", "Document Images Required", "Website Path", "Title Line 1", \
         "Title Line 2", "Title Line 3", "Title Line 4", "Country to remove 1", "Country to remove 2", \
         "Person1", "Person2", "Person3", "Person4"]
fill = [36, 34, 29, 19, 21, 25, 7, 29, 31, 31, 31, 31, 16, 16, 35, 35, 35, 35]
root = ""
param0 = param1 = param2 = param3 = param4 = param5 = param6 = param7 = param8 = param9 = ""
param10 = param11 = param12 = param13 = param14 = param15 = param16 = param17 = param18 = ""
label0 = label1 = label2 = label3 = label4 = label5 = label6 = label7 = label8 = label9 = ""
label10 = label11 = label12 = label13 = label14 = label15 = label16 = label17 = ""

def get_params():
    global ged_file_name, initial_family, title_page_required, title1, title2, title3, title4, \
           contents_required, index_required, page_size, document_images_required, website_path, \
           country_to_remove1, country_to_remove2, person1, person2, person3, person4

    paramfile = open('params.txt','r')
    while True:
        s = paramfile.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        param_name = x[0]
        param_value = x[1]
        if param_name == field[0]: ged_file_name = param_value
        if param_name == field[1]: page_size = param_value
        if param_name == field[2]: initial_family = int(param_value)
        if param_name == field[3]: title_page_required = param_value
        if param_name == field[4]: contents_required = param_value
        if param_name == field[5]: index_required = param_value
        if param_name == field[6]: document_images_required = param_value
        if param_name == field[7]: website_path = param_value
        if param_name == field[8]: title1 = param_value
        if param_name == field[9]: title2 = param_value
        if param_name == field[10]: title3 = param_value
        if param_name == field[11]: title4 = param_value
        if param_name == field[12]: country_to_remove1 = param_value
        if param_name == field[13]: country_to_remove2 = param_value
        if param_name == field[14]: person1 = int(param_value)
        if param_name == field[15]: person2 = int(param_value)
        if param_name == field[16]: person3 = int(param_value)
        if param_name == field[17]: person4 = int(param_value)

def save_params():
    paramfile = open('params.txt','w')
    paramfile.write(field[0]  + '~' + param0.get()+ '\n')
    paramfile.write(field[1]  + '~' + param1.get() + '\n')
    paramfile.write(field[2]  + '~' + param2.get() + '\n')
    paramfile.write(field[3]  + '~' + param3.get() + '\n')
    paramfile.write(field[4]  + '~' + param4.get() + '\n')
    paramfile.write(field[5]  + '~' + param5.get() + '\n')
    paramfile.write(field[6]  + '~' + param6.get() + '\n')
    paramfile.write(field[7]  + '~' + param7.get() + '\n')
    paramfile.write(field[8]  + '~' + param8.get() + '\n')
    paramfile.write(field[9]  + '~' + param9.get() + '\n')
    paramfile.write(field[10] + '~' + param10.get() + '\n')
    paramfile.write(field[11] + '~' + param11.get() + '\n')
    paramfile.write(field[12] + '~' + param12.get() + '\n')
    paramfile.write(field[13] + '~' + param13.get() + '\n')
    paramfile.write(field[14] + '~' + param14.get() + '\n')
    paramfile.write(field[15] + '~' + param15.get() + '\n')
    paramfile.write(field[16] + '~' + param16.get() + '\n')
    paramfile.write(field[17] + '~' + param17.get() + '\n')
    paramfile.close()
    
want_ui = True
    
def save():
    save_params()
    root.destroy()

def exit():
    root.destroy()

def display_param(n, l, p, t):
    f = field[n]+spaces[0:fill[n]]
    l = Label(root, text=f,font=(font, font_size), width=label_width)
    l.grid(row=n+1, column=1, sticky=W)
    p = Entry(root, font=(font, font_size), width=field_width)
    p.insert(0, t)
    p.grid(row=n+1, column=2, sticky=W)
    return(p)

def edit_params():
    global root, param0, param1, param2, param3, param4, param5, param6, param7, param8, param9, param10,\
           param11, param12, param13, param14, param15, param16, param17

    get_params()
    
    root = Tk()
    root.title('Parameters')
    root.geometry('1400x700') #width, height
    
    space1 = Label(root, text='', font=(font, font_size))
    space1.grid(row=0, column=0)

    param0  = display_param(0, label0, param0, ged_file_name)
    param1  = display_param(1, label1, param1, page_size)
    param2  = display_param(2, label2, param2, initial_family)
    param3  = display_param(3, label3, param3, title_page_required)
    param4  = display_param(4, label4, param4, contents_required)
    param5  = display_param(5, label5, param5, index_required)
    param6  = display_param(6, label6, param6, document_images_required)
    param7  = display_param(7, label7, param7, website_path)
    param8  = display_param(8, label8, param8, title1)
    param9  = display_param(9, label9, param9, title2)
    param10 = display_param(10, label10, param10, title3)
    param11 = display_param(11, label11, param11, title4)
    param12 = display_param(12, label12, param12, country_to_remove1)
    param13 = display_param(13, label13, param13, country_to_remove2)
    param14 = display_param(14, label14, param14, person1)
    param15 = display_param(15, label15, param15, person2)
    param16 = display_param(16, label16, param16, person3)
    param17 = display_param(17, label17, param17, person4)

    space2 = Label(root, text='', font=(font, font_size))
    space2.grid(row=19, column=0)

    button1 = Button(root, text='Quit', font=(font, font_size), command=exit, width=10)
    button1.grid(row=20, column=1)
    button2 = Button(root, text='Save', font=(font, font_size), command=save, width=10)
    button2.grid(row=20, column=2, sticky=W)

    root.mainloop()
    
#edit_params()
