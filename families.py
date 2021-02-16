
class families_class(object):
    def __init__(self, family_id=None, husband_id=None, wife_id=None, marriage_date=None, marriage_place=None):
        self.family_id = family_id
        self.husband_id = husband_id
        self.wife_id = wife_id
        self.marriage_date = marriage_date
        self.marriage_place = marriage_place

families = []

column_heading = ["FamilyID", "Husband", "Wife", "Marriage Date", "Marriage Place" ]

def read_families():
    file = open('families.txt','r')
    families.clear()
    s = file.readline()
    add_family(0, 0, 0, "", "")
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_family(int(x[0]), int(x[1]), int(x[2]),  x[3],  x[4])

    print("Read " + str(len(families)-2) + " families")
    
def add_family(family_id, husband_id, wife_id, marriage_date, marriage_place):
    families.append(families_class(family_id, husband_id, wife_id, marriage_date, marriage_place))

def write_families():
    file = open('families.txt','w')
    line = ""
    for i in range(len(column_heading)):
        line = line + column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(families)):
        line = str(families[i].family_id) + "~"
        line = line + str(families[i].husband_id) + "~"
        line = line + str(families[i].wife_id) + "~"
        line = line + families[i].marriage_date + "~"
        line = line + families[i].marriage_place
        file.write(line + '\n')

    file.close()
    print("Written " + str(len(families)) + " families")

def get_father_id(id):
    family_id = ind.get_family_where_child(id)
    if family_id != 0:
#        i = int(f)
        return families[family_id].husband_id
    else:
        return 0

def get_mother_id(id):
    family_id = ind.get_family_where_child(id)
    if family_id != "":
#        i = int(f)
        return families[family_id].wife_id
    else:
        return 0

