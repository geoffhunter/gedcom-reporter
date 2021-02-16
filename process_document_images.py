import os
import params

class file_type_order_class(object):
    def __init__(self, file_type=None, order=None):
        self.file_type = file_type
        self.order = order

def process_document_images():
    print ("\nProcessing document images")
    params.get_params()
    font_style = "style=font-family:\"courier\";font-weight:bold;"

    file_type_order = []
    ftofile = open("File_Type_Order.txt","r")
    c_fto = 0
    
    while True:
        s = ftofile.readline()
        s = s.strip()
        if s == "":
            break
        file_type = s[0:s.find("~")]
        order = s[s.find("~")+1:]
        file_type_order.append(file_type_order_class(file_type, order))
        c_fto = c_fto + 1
    
    # get list of records to sort
    records = []
    last_person_id = ""
    ifile = ""
    c_records = 0
    for r, d, f in os.walk(params.website_path + "/images"):
        for file in f:
            if '.jpg' in file or '.JPG' in file:
                file_name = file[0:len(file)-4]
                person_id = file_name[0:file_name.find("-") - 1]
                file_type = file_name[file_name.find("-") + 2:]
                for i in range(0,c_fto):
                    if file_type_order[i].file_type == file_type:
                        order = file_type_order[i].order
                        break;
                
                record = person_id + "~" + order + "~" + file_type
                records.append(record)
                c_records = c_records + 1
    
    # sort records
    records.sort()
    
    # process sorted records
    for i in range(0,c_records):
        record = records[i]
        person_id = record[0:record.find("~")]
        record = record[record.find("~") + 1:]
        file_type = record[record.find("~")+1:]
        name_and_type = person_id + " - " + file_type
        
        if person_id != last_person_id:
            if last_person_id != "":
                ifile.write("</body>")
                ifile.write("</html>")
                ifile.close()

            file_name_index = params.website_path + "/Index " + person_id + ".html"
            ifile = open(file_name_index,"w")
            ifile.write("<!DOCTYPE html PUBLIC \" -//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">\n")
            ifile.write("<html>\n")
            ifile.write("<head>\n")
            ifile.write("<meta content=\"text/html; charset=ISO-8859-1\" http-equiv=\"content-type\">\n")
            ifile.write("<title>Contents</title>\n")
            ifile.write("</head>\n")
            ifile.write("<body " + font_style + ">\n")
            ifile.write("<p class=\"font_8\">" + person_id + "</p>\n")
        
        long_file_type = file_type
        long_file_type = long_file_type.replace("BIR", "Birth Index Record")
        long_file_type = long_file_type.replace("MIR", "Marriage Index Record")
        long_file_type = long_file_type.replace("DIR", "Death Index Record")
        file_name_HTML = "Record " + name_and_type + ".html"
        ifile.write("<p style=\"margin-left: 40px;\" class=\"font_8\">\n")
        ifile.write("<span style=\"text-decoration: underline;\"><a href=\"" + file_name_HTML + "\">" + long_file_type + "</a></span></p>\n")
            
        file_name_HTML = params.website_path + "\\" + file_name_HTML
        hfile = open(file_name_HTML,"w")
        hfile.write("<!DOCTYPE html PUBLIC \" -//W3C//DTD HTML 4.01//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">\n")
        hfile.write("<html>\n")
        hfile.write("<head>\n")
        hfile.write("<meta content=\"text/html; charset=ISO-8859-1\" http-equiv=\"content-type\">\n")
        hfile.write("<title>" + name_and_type + "</title>\n")
        hfile.write("</head>\n")
        hfile.write("<body " + font_style + ">\n")
        hfile.write("<img style=\"width: 100%; height: 100%;\" alt=\"" + name_and_type + "\" src=\"images\\" + name_and_type + ".jpg\">\n")
        hfile.write("</body>\n")
        hfile.write("</html>\n")
        hfile.close()
        last_person_id = person_id
        
    if ifile != "":
        ifile.write("<body " + font_style + ">")
        ifile.write("</html>")
        ifile.close()
            
#process_document_images()