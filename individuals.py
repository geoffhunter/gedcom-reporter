
class individuals_class(object):
    def __init__(self, individual_id=None, surname=None, forename=None, name = None, \
                 birth_date=None, birth_place=None, baptism_date=None, baptism_place=None, \
                 marriage_date=None, marriage_place=None, death_date=None, death_place=None, \
                 burial_date=None, burial_place=None, family_where_child=None, family_where_spouse=None):
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
        
individuals = []

column_heading = ["IndividualID", "Surname", "Forename(s)", "Name", "Birth Date", "Birth Place", \
                  "Baptism Date", "Baptism Place", "Marriage Date", "Marriage Place", "Death Date", \
                  "Death Place", "Burial Date", "Burial Place", "FamilyWhereChild", "FamilyWhereSpouse"]

def read_individuals():
    file = open('individuals.txt','r')
    individuals.clear()
    s = file.readline()
    i = 0
    add_individual("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_individual(int(x[0]), x[1], x[2],  x[3],  x[4],  x[5],  x[6],  x[7], \
                       x[8], x[9], x[10], x[11], x[12], x[13], int(x[14]), int(x[15]))
        i = i + 1

    print("Read " + str(len(individuals)-2) + " individuals")

def add_individual( \
        individual_id, surname, forename, name, birth_date, birth_place, \
        baptism_date, baptism_place, marriage_date, marriage_place, death_date, death_place, \
        burial_date, burial_place, family_where_child, family_where_spouse):
    individuals.append(individuals_class( \
        individual_id, surname, forename, name, birth_date, birth_place, \
        baptism_date, baptism_place, marriage_date, marriage_place, death_date, death_place,  \
        burial_date, burial_place, family_where_child, family_where_spouse))

def write_individuals():
    file = open('individuals.txt','w')
    line = ""
    for i in range(len(column_heading)):
        line = line + column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(individuals)):
        line = str(individuals[i].individual_id) + "~"
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
        line = line + str(individuals[i].family_where_spouse)
        file.write(line + '\n')

    file.close()

    print("Written " + str(len(individuals)) + " individuals")

def get_person_name(id):
    i = int(id)
    n = individuals[i].surname + ", " + individuals[i].forename
    return n.strip()

def get_birth_date(id):
    i = int(id)
    bd = individuals[i].birth_date
    bd = bd[-4:]
    return "b. " + bd

def get_family_where_child(id):
    i = int(id)
    return individuals[i].family_where_child

#read_individuals()
