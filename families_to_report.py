#import xlrd
#import xlsxwriter
import individuals as ind
import families as fam

class families_to_report_class(object):
    def __init__(self, family_number=None, family_id=None, to_report=None):
        self.family_number = family_number
        self.family_id = family_id
        self.to_report = to_report

families_to_report = []

def read_families_to_report():
    file = open('FamiliesToReport.txt','r')
    families_to_report.clear()
    while True:
        s = file.readline()
        s = s.strip()
        if s == '':
            break
        x = s.split("~")
        add_family_to_report(int(x[0]), int(x[1]), x[6])

#    print("Read " + str(len(families_to_report)) + " families to report")
#    workbook = xlrd.open_workbook("FamiliesToReport.xlsx") 
#    sheet = workbook.sheet_by_index(0) 
#    families_to_report.clear()
#    i = 0
#    while True:
#        if i == sheet.nrows: break
#        add_family_to_report(sheet.cell_value(i, 0), sheet.cell_value(i, 1), sheet.cell_value(i, 6))
#        i = i + 1

def add_family_to_report(family_number, family_id, to_report):
    families_to_report.append(families_to_report_class(family_number, family_id, to_report))
    
def write_families_to_report():
#    workbook = xlsxwriter.Workbook('FamiliesToReport.xlsx')
#    worksheet = workbook.add_worksheet()
    file = open('FamiliesToReport.txt','w')
    for i in range(0,len(families_to_report)):
        family_id = int(families_to_report[i].family_id)
        husbands_name = ""
        husbands_birth_date = ""
        if fam.families[family_id].husband_id != "":
            husband_id = int(fam.families[family_id].husband_id)
            if husband_id > 0:
                husbands_name = ind.get_person_name(husband_id)
                husbands_birth_date = ind.get_birth_date(husband_id)

        wifes_name = ""
        wifes_birth_date = ""
        if fam.families[family_id].wife_id != "":
            wife_id = int(fam.families[family_id].wife_id)
            if wife_id > 0:
                wifes_name = ind.get_person_name(wife_id)
                wifes_birth_date = ind.get_birth_date(wife_id)
#        n = int(family_id)
#        worksheet.write(i, 0,  families_to_report[i].family_number)
#        worksheet.write(i, 1,  n)
#        worksheet.write(i, 2,  husbands_name)
#        worksheet.write(i, 3,  husbands_birth_date)
#        worksheet.write(i, 4,  wifes_name)
#        worksheet.write(i, 5,  wifes_birth_date)
#        worksheet.write(i, 6,  families_to_report[i].to_report)
        line = str(families_to_report[i].family_number) + "~"
#        n = families_to_report[i].family_id
        s = str(family_id)
        line = line + s + "~"
        line = line + husbands_name + "~"
        line = line + husbands_birth_date + "~"
        line = line + wifes_name + "~"
        line = line + wifes_birth_date + "~"
        line = line + families_to_report[i].to_report
        file.write(line + '\n')
    
    file.close()
    #    workbook.close()
  
read_families_to_report()

