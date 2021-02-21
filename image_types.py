image_types = []

def read_image_types():
    global image_types

    imagetypefile = open("ImageTypes.txt","r")
    while True:
        s = imagetypefile.readline()
        s = s.strip()
        if s == "":
            break
        image_types.append(s)

