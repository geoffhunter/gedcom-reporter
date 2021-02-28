# gedcom-reporter

Allows the user to produce an ancestry report, in Microsoft Word (.docx) format, from a GEDCOM format file. 

Note: this utility needs the .py and .txt files in the gedcom-file-processor repository to function

The report shows each family on a new page, starting with a specified family. 

Each page shows a family reference code, then the husband, wife and children with birth, baptism, marriage, death and burial places and dates. 

Names are followed by a family reference code indicating where the person appears elsewhere in the report, to assist in moving around the report. Names
may also have a hyperlink to the page where they also also appear as a parent or child. 

Each page is followed by further pages showing the husband's family and ancestors, where available, then the wife's family and ancestors, where available.

There are also options for a Title page, showing the name of the starting person, and Contents and Index pages. For neatness, the 'home' country, 
e.g. England and United Kingdom may be removed from place names, or retained if you wish. The 'home' countries to be removed can be specified.
