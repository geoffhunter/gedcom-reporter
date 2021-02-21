class index_class(object):
    def __init__(self, key=None, surname=None, forename=None, birth_year=None, family_number_where_spouse=None, \
                 family_number_where_child=None):
        self.key = key
        self.surname = surname
        self.forename = forename
        self.birth_year = birth_year
        self.family_number_where_spouse = family_number_where_spouse
        self.family_number_where_child = family_number_where_child
        
index = []

def add_index(key, surname, forename, birth_year, family_number_where_spouse, family_number_where_child):
    index.append(index_class(key, surname, forename, birth_year, family_number_where_spouse,
                             family_number_where_child))
    
