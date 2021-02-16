import params as params
import individuals as ind
import families as fam
import children as chi

def process_ged_file():
    print("Processing GED File")
    params.get_params()
    fam.c_families = 0
    chi.c_children = 0
    
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
    
    family_id = 0
    husband_id = 0
    wife_id = 0
    
    ind.individuals.clear()
    ind.add_individual("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    fam.families.clear()
    fam.add_family("", "", "", "", "")
    chi.children.clear()
    chi.add_child("", "")
    
    child_id = 0

    gfile = open(params.ged_file_name, "r")

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
                    ind.add_individual(individual_id, surname, forename, name, birth_date, birth_place, \
                                       baptism_date, baptism_place, marriage_date, marriage_place,\
                                       death_date, death_place, burial_date, burial_place, \
                                       family_where_child, family_where_spouse)
                    have_record = False
                if family_id != 0:
                    fam.add_family(family_id, husband_id, wife_id, marriage_date, marriage_place)
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
#                print(s, husband, husband_id)
                if ind.individuals[husband_id].family_where_spouse == 0:
                    ind.individuals[husband_id].family_where_spouse = family_id
            if tag == "WIFE":
                wife = s[9:]
                wife = wife[0:len(wife)-1]
                wife_id = int(wife)
                if ind.individuals[wife_id].family_where_spouse == 0:
                    ind.individuals[wife_id].family_where_spouse = family_id
            if tag == "CHIL":
                child = s[9:]
                child = child[0:len(child)-1]
                child_id = int(child)
                ind.individuals[child_id].family_where_child = family_id
                chi.add_child(family_id, child_id)
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

    ind.write_individuals()
    fam.write_families()
    chi.write_children()
    
process_ged_file()
