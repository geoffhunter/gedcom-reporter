
class children_class(object):
    def __init__(self, family_id=None, child_id=None):
        self.family_id = family_id
        self.child_id = child_id

children = []

column_heading = ["FamilyID", "ChildID"]

def read_children():
    file = open('children.txt','r')
    children.clear()
    s = file.readline()
    add_child(0, 0)
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_child(int(x[0]), int(x[1]))

    print("Read " + str(len(children)-2) + " children")
    
def add_child(family_id, child_id):
    children.append(children_class(family_id, child_id))

def write_children():
    file = open('children.txt','w')
    line = ""
    for i in range(len(column_heading)):
        line = line + column_heading[i]  + '~'
    line = line[0:len(line)-1]        
    file.write(line + '\n')

    for i in range(1,len(children)):
        line = str(children[i].family_id) + "~"
        line = line + str(children[i].child_id)
        file.write(line + '\n')

    file.close()
    print("Written " + str(len(children)) + " children")
