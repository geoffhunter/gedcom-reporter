page_size = ""
font_name = ""
page_height = 0
page_width = 0
left_margin = 0
right_margin = 0
top_margin = 0
bottom_margin = 0
header = 0
footer = 0
title_font_size = 0
title_page_rows = 0
normal_font_size = 0
index_rows = 0
max_image_height = 0
max_image_width = 0
max_image_ratio = 0

def define_page_size(size):
    global page_size, font_name, page_height, page_width, left_margin, right_margin, top_margin, bottom_margin, \
           header, footer, title_font_size, title_page_rows, normal_font_size, index_rows, \
           max_image_height, max_image_width, max_image_ratio
    
    page_size = size
    font_name = "Calibri"
    if page_size == "A4":
        page_height = 297
        page_width = 210
        left_margin = 15
        right_margin = 15
        top_margin = 15
        bottom_margin = 15
        header = 10
        footer = 10
        title_font_size = 26
        title_page_rows = 26
        normal_font_size = 11
        index_rows = 28
        max_image_height = page_height - top_margin - bottom_margin - header - footer - 10
        max_image_width = page_width - left_margin - right_margin - 10
        max_image_ratio = max_image_height / max_image_width
        return(1)
    if page_size == "A5":
        page_height = 210
        page_width = 148
        left_margin = 10
        right_margin = 10
        top_margin = 10
        bottom_margin = 10
        header = 6
        footer = 6
        title_font_size = 18
        title_page_rows = 18
        normal_font_size = 8
        index_rows = 20
        max_image_height = page_height - top_margin - bottom_margin - header - footer - 10
        max_image_width = page_width - left_margin - right_margin - 10
        max_image_ratio = max_image_height / max_image_width
        return(1)
    print ("Unsupported Page Size")
    return(0)
        
