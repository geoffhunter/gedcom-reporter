from tkinter import *

def process_ged_file():
#    print("Processing GED File")
    get_params()
    
    individual_id = 0
    surname = ""
    forename = ""
    name = ""
    birth_date = ""
    baptism_date = ""
    marriage_date = ""
    death_date = ""
    burial_date = ""
    birth_place = ""
    baptism_place = ""
    marriage_place = ""
    death_place = ""
    burial_place = ""
    family_where_child = ""
    family_where_spouse = ""
    sex = ""
    
    family_id = 0
    husband_id = 0
    wife_id = 0
    
    individuals.clear()
    add_individual("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    families.clear()
    add_family("", "", "", "", "", "")
    children.clear()
    add_child("", "", "")
    
    child_id = 0

    gfile = open(ged_file_name, "r")

    have_record = False
    while True:
        line = gfile.readline()
        s = line.strip()
        if s == "":
            break
        
        c1 = s[:1]
        if c1 == "0":
            if have_record:
                if individual_id != 0:
                    name = forename
                    if name != "": name = name + " "
                    name = name + surname
                    add_individual("N", individual_id, surname, forename, name, birth_date, birth_place, \
                                       baptism_date, baptism_place, marriage_date, marriage_place,\
                                       death_date, death_place, burial_date, burial_place, \
                                       family_where_child, family_where_spouse, sex)
                    have_record = False
                if family_id != 0:
                    add_family("N", family_id, husband_id, wife_id, marriage_date, marriage_place)
                    have_record = False
            record_type = s[-4:]
            if record_type == "INDI" or record_type == " FAM":
                individual_id = 0
                surname = ""
                forename = ""
                birth_date = ""
                baptism_date = ""
                marriage_date = ""
                death_date = ""
                burial_date = ""
                birth_place = ""
                baptism_place = ""
                marriage_place = ""
                death_place = ""
                burial_place = ""
                family_where_child = 0
                family_where_spouse = 0
                family_id = 0
                husband_id = 0
                wife_id = 0
                child_id = 0
                sex = ""
                tag = ""
            else:
                if record_type != "HEAD" and record_type != "SOUR" and record_type != "REPO" and record_type != "TRLR":
                    print ("Unknown", c1, record_type)
                    print (s)
                
            if record_type == "INDI":
                individual = s[4:]
                individual = individual[0:individual.find("@")]
                individual_id = int(individual)
                have_record = True
            if record_type == " FAM":
                family = s[4:]
                family = family[0:family.find("@")]
                family_id = int(family)
                have_record = True
        if c1 == "1":
            tag = s[2:6]
            if tag == "NAME":
                if surname == "" and forename == "":
                    individual_name = s[7:]
                    p = individual_name.find("/")
                    if (p >= 0):
                        if p == 0:
                            forename = ""
                        else:
                            forename = individual_name[0:p-1]
                        surname = individual_name[p+1:]
                        p = surname.find("/")
                        if (p > 0):
                            surname = surname[0:p]
                    else:
                        forename = individual_name
            if tag == "BIRT":
                date_place_type = tag
            if tag == "BAPM":
                date_place_type = tag
            if tag == "MARR":
                date_place_type = tag
            if tag == "DEAT":
                date_place_type = tag
            if tag == "BURI":
                date_place_type = tag
            if tag == "RESI":
                date_place_type = tag
            if tag == "EVEN":
                date_place_type = tag
            if tag == "HUSB":
                husband = s[9:]
                husband = husband[0:len(husband)-1]
                husband_id = int(husband)
                if individuals[husband_id].family_where_spouse == 0:
                    individuals[husband_id].family_where_spouse = family_id
            if tag == "WIFE":
                wife = s[9:]
                wife = wife[0:len(wife)-1]
                wife_id = int(wife)
                if individuals[wife_id].family_where_spouse == 0:
                    individuals[wife_id].family_where_spouse = family_id
            if tag == "CHIL":
                child = s[9:]
                child = child[0:len(child)-1]
                child_id = int(child)
                individuals[child_id].family_where_child = family_id
                add_child(family_id, child_id, "")
            if tag == "SEX ":
                sex = s[6:]
        if c1 == "2":
            tag = s[2:6]
            if tag == "DATE":
                s_date = s[7:]
                if date_place_type == "BIRT":
                    birth_date = s_date
                if date_place_type == "BAPM":
                    baptism_date = s_date
                if date_place_type == "MARR":
                    marriage_date = s_date
                if date_place_type == "DEAT":
                    death_date = s_date
                if date_place_type == "BURI":
                    burial_date = s_date
            if tag == "PLAC":
                place = s[7:]
                if place[-1:] == ".":
                    place = place[0:len(place)-1]
                if date_place_type == "BIRT":
                    birth_place = place
                if date_place_type == "BAPM":
                    baptism_place = place
                if date_place_type == "MARR":
                    marriage_place = place
                if date_place_type == "DEAT":
                    death_place = place
                if date_place_type == "BURI":
                    burial_place = place
    gfile.close()

    write_individuals()
    write_families()
    write_children()

class individuals_class(object):
    def __init__(self, tag=None, individual_id=None, surname=None, forename=None, name = None, \
                 birth_date=None, birth_place=None, baptism_date=None, baptism_place=None, \
                 marriage_date=None, marriage_place=None, death_date=None, death_place=None, \
                 burial_date=None, burial_place=None, family_where_child=None, family_where_spouse=None, \
                 sex=None):
        self.tag = tag
        self.individual_id = individual_id
        self.surname = surname
        self.forename = forename
        self.name = name
        self.birth_date = birth_date
        self.birth_place = birth_place
        self.baptism_date = baptism_date
        self.baptism_place = baptism_place
        self.marriage_date = marriage_date
        self.marriage_place = marriage_place
        self.death_date = death_date
        self.death_place = death_place
        self.burial_date = burial_date
        self.burial_place = burial_place
        self.family_where_child = family_where_child
        self.family_where_spouse = family_where_spouse
        self.sex = sex
        
individuals = []

ind_column_heading = ["Tag", "IndividualID", "Surname", "Forename(s)", "Name", "Birth Date", "Birth Place", \
                  "Baptism Date", "Baptism Place", "Marriage Date", "Marriage Place", "Death Date", \
                  "Death Place", "Burial Date", "Burial Place", "FamilyWhereChild", "FamilyWhereSpouse", \
                  "Sex"]

def read_individuals():
    file = open('individuals.txt','r')
    individuals.clear()
    s = file.readline()
    i = 0
    add_individual("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_individual(x[0], int(x[1]), x[2],  x[3],  x[4],  x[5],  x[6],  x[7], \
                       x[8], x[9], x[10], x[11], x[12], x[13], x[14], int(x[15]), int(x[16]), x[17])
        i = i + 1

#    print("Read " + str(len(individuals)-1) + " individuals")

def add_individual( \
        tag, individual_id, surname, forename, name, birth_date, birth_place, \
        baptism_date, baptism_place, marriage_date, marriage_place, death_date, death_place, \
        burial_date, burial_place, family_where_child, family_where_spouse, sex):
    individuals.append(individuals_class( \
        tag, individual_id, surname, forename, name, birth_date, birth_place, \
        baptism_date, baptism_place, marriage_date, marriage_place, death_date, death_place,  \
        burial_date, burial_place, family_where_child, family_where_spouse, sex))

def write_individuals():
    file = open('individuals.txt','w')
    line = ""
    for i in range(len(ind_column_heading)):
        line = line + ind_column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(individuals)):
        line = individuals[i].tag + "~"
        line = line + str(individuals[i].individual_id) + "~"
        line = line + individuals[i].surname + "~"
        line = line + individuals[i].forename + "~"
        line = line + individuals[i].name + "~"
        line = line + individuals[i].birth_date + "~"
        line = line + individuals[i].birth_place + "~"
        line = line + individuals[i].baptism_date + "~"
        line = line + individuals[i].baptism_place + "~"
        line = line + individuals[i].marriage_date + "~"
        line = line + individuals[i].marriage_place + "~"
        line = line + individuals[i].death_date + "~"
        line = line + individuals[i].death_place + "~"
        line = line + individuals[i].burial_date + "~"
        line = line + individuals[i].burial_place + "~"
        line = line + str(individuals[i].family_where_child) + "~"
        line = line + str(individuals[i].family_where_spouse) + "~"
        line = line + individuals[i].sex
        file.write(line + '\n')

    file.close()

#    print("Written " + str(len(individuals)-1) + " individuals")

def get_person_name(id):
    n = individuals[id].surname + ", " + individuals[id].forename
    return n.strip()

def get_birth_year(id):
    bd = individuals[id].birth_date
    bd = bd[-4:]
    return "b. " + bd

def get_family_where_child(id):
    return individuals[id].family_where_child

def get_name(i):
    name = individuals[i].forename
    surname = individuals[i].surname
    if name != '' and surname != '':
        name = name + ' '
    name = name + surname.upper()
    return(name)
    
class families_class(object):
    def __init__(self, tag=None, family_id=None, husband_id=None, wife_id=None, marriage_date=None, \
                 marriage_place=None):
        self.tag = tag
        self.family_id = family_id
        self.husband_id = husband_id
        self.wife_id = wife_id
        self.marriage_date = marriage_date
        self.marriage_place = marriage_place

families = []

fam_column_heading = ["Tag", "FamilyID", "Husband", "Wife", "Marriage Date", "Marriage Place"]

def read_families():
    file = open('families.txt','r')
    families.clear()
    s = file.readline()
    add_family("", 0, 0, 0, "", "")
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_family(x[0], int(x[1]), int(x[2]), int(x[3]), x[4], x[5])

#    print("Read " + str(len(families)-1) + " families")
    
def add_family(tag, family_id, husband_id, wife_id, marriage_date, marriage_place):
    families.append(families_class(tag, family_id, husband_id, wife_id, marriage_date, marriage_place))

def write_families():
    file = open('families.txt','w')
    line = ""
    for i in range(len(fam_column_heading)):
        line = line + fam_column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(families)):
        line = families[i].tag + "~"
        line = line + str(families[i].family_id) + "~"
        line = line + str(families[i].husband_id) + "~"
        line = line + str(families[i].wife_id) + "~"
        line = line + families[i].marriage_date + "~"
        line = line + families[i].marriage_place
        file.write(line + '\n')

    file.close()
#    print("Written " + str(len(families)-1) + " families")

def get_father_id(id):
    family_id = get_family_where_child(id)
    if family_id != 0:
        return families[family_id].husband_id
    else:
        return 0

def get_mother_id(id):
    family_id = get_family_where_child(id)
    if family_id != "":
        return families[family_id].wife_id
    else:
        return 0

class children_class(object):
    def __init__(self, family_id=None, child_id=None, tag=None):
        self.family_id = family_id
        self.child_id = child_id
        self.tag = tag

children = []

chi_column_heading = ["FamilyID", "ChildID", "Tag"]

def read_children():
    file = open('children.txt','r')
    children.clear()
    s = file.readline()
    add_child(0, 0, "")
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_child(int(x[0]), int(x[1]), "")

#    print("Read " + str(len(children)-1) + " children")
    
def add_child(family_id, child_id, tag):
    children.append(children_class(family_id, child_id, tag))

def write_children():
    file = open('children.txt','w')
    line = ""
    for i in range(len(chi_column_heading)):
        line = line + chi_column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(children)):
        line = str(children[i].family_id) + "~"
        line = line + str(children[i].child_id) + "~"
        line = line + str(children[i].tag)
        file.write(line + '\n')

    file.close()
#    print("Written " + str(len(children)-1) + " children")

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
    
process_ged_file()
