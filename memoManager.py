import os
import shutil
import subprocess
import random
import string
import xlwings as xw
import textract
import glob

def copyMemo( number, staff, role, password ) :
    if os.path.isdir("Allstaff/" + staff )==False : os.mkdir("Allstaff/" + staff )
    lockpwd = ''.join(random.choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for _ in range(20))
    cmd = ['pdftk', 'Memos/' + number + '.pdf', 'output', 'Allstaff/' + staff + '/' + role + '_' + number + '.pdf', 'owner_pw', lockpwd, 'user_pw', password ]
    proc = subprocess.Popen(cmd)
    proc.communicate()


# Make a dictionary from excel containing the students
wb = xw.Workbook( '/Users/gareth/Desktop/DS/DS-list-2017-2018-maths.xlsx')
n, student_dictionary, data = 0, {}, xw.Range('Sheet1', 'B3', wkb=wb).table.value
for col in data[0] :
    tdict = {}
    tdict["advisor"], tdict["personal"], tdict["modules"] = data[3][n], data[4][n], []
    for mod in [9,11,13,15,17,19] :
        if data[mod][n]!="End" : tdict["modules"].append( data[mod][n] ) 
    student_dictionary[ str(int(col)) ] = tdict 
    n+=1 

#print( student_dictionary )

# Get data on staff numbers from excel
staffdict, staffdata = {}, xw.Range("Sheet4","A2").table.value
for row in staffdata : staffdict[ row[1] ] = "Sn@" + str( int(row[0]) )

# Make a dictionary of who teaches each module
n, modteach, modules, teachers = 0, {}, xw.Range("Sheet3","A1:A83", wkb=wb).value, xw.Range("Sheet3","B1:B83", wkb=wb).value
for module in modules : 
    modteach[module] = teachers[n].split()
    # Check we have staff numbers for everyone who teaches a module
    for staff in modteach[module] :
        if staff not in staffdict : RuntimeError("staff number not found for " + staff)
    n+=1 

# Now work through all the memos that were downloaded
os.mkdir("Memos")
os.mkdir("Allstaff")
for memo in os.listdir("/Users/gareth/Desktop/DS/Downloadedmemos/") :
    if memo.endswith(".docx") :
       # Get the student number from the memo
       text = textract.process('/Users/gareth/Desktop/DS/Downloadedmemos/' + memo ).decode("utf-8")
       start = text.find("Student No:")
       end = text.find("Course:")
       studentno = text[start:end].replace("Student No:","").rstrip().strip()

       # Now find the details of the student from excel
       if studentno not in student_dictionary : RuntimeError("Could not find student number " + studentno + " for memo " + memo )
       thisstudent = student_dictionary[ studentno ] 

       # Now copy the memo to my stash of memos and rename to student number
       shutil.copy( '/Users/gareth/Desktop/DS/Downloadedmemos/' + memo.replace(".docx",".pdf"), "Memos/" + studentno + ".pdf" )
       # Copy the students memo for the advisor of study 
       copyMemo( studentno, thisstudent["advisor"], "advisor", staffdict[thisstudent["advisor"]] )
       # Copy the students memo for the personal tutor
       copyMemo( studentno, thisstudent["personal"], "personal", staffdict[thisstudent["personal"]] )
       # Now make copies for modules
       for module in thisstudent["modules"] :
           for staff in modteach[module] : copyMemo( studentno, staff, module, staffdict[staff] )
