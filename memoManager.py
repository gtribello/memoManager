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
wbp = xw.Workbook( '/Users/gareth/Desktop/DS/DS-list-2017-2018-physics.xlsm')
# Get data on staff numbers from exccel
n, staffdict, staffdata, staffnumbers = 0, {}, xw.Range("options","G3:G71").value, xw.Range("options","F3:F71").value
for name in staffdata : 
    staffdict[ name ] = "Sn@" + str( int(staffnumbers[n]) )
    n += 1

#for k,n in staffdict.items() : print( k, n )

# Make a dictionary of who teaches each module
n, modteach, modules, teachers = 0, {}, xw.Range("options","K3:K86", wkb=wbp).value, xw.Range("options","L3:L86", wkb=wbp).value
for module in modules : 
    modteach[module] = teachers[n].split()
    # Check we have staff numbers for everyone who teaches a module
    for staff in modteach[module] :
        if staff not in staffdict : RuntimeError("staff number not found for " + staff)
    n+=1 

#for k, n in modteach.items() : print( k, n )

n, student_dictionary, data = 0, {}, xw.Range('main', 'B8', wkb=wbp).table.value
for col in data[0] :
    tdict = {}
    tdict["modules"] = []
    if (data[12][n]!=None) & (data[12][n]!="End") : tdict["advisor"] = data[12][n] 
    if (data[13][n]!=None) & (data[13][n]!="End") : tdict["personal"] = data[13][n]
    if (data[36][n]!=None) & (data[36][n]!="End") : tdict["supervisor"] = data[36][n] 
    for mod in [16,18,20,22,24,26,28,30,32,34] :
        if (data[mod][n]!=None) & (data[mod][n]!="End") : tdict["modules"].append( data[mod][n] )
    student_dictionary[ str(int(col)) ] = tdict 
    n += 1

wbm = xw.Workbook( '/Users/gareth/Desktop/DS/DS-list-2017-2018-maths.xlsx')
n, data = 0, xw.Range('Sheet1', 'B3', wkb=wbm).table.value
for col in data[0] :
    tdict = {}
    if (data[3][n]!=None) & (data[3][n]!="End") : tdict["advisor"] = data[3][n] 
    if (data[4][n]!=None) & (data[4][n]!="End") : tdict["personal"] = data[4][n]
    if (data[23][n]!=None) & (data[23][n]!="End") : tdict["supervisor"] = data[23][n]
    tdict["modules"] = []
    for mod in [9,11,13,15,17,19,21] :
        if (data[mod][n]!=None) & (data[mod][n]!="End") : tdict["modules"].append( data[mod][n] ) 
    student_dictionary[ str(int(col)) ] = tdict 
    n+=1 

#for k, n in student_dictionary.items() : print( k, n )


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
       print("Making memos for student number " + studentno + " with memo " + memo )
       if studentno not in student_dictionary : RuntimeError("Could not find student number " + studentno + " for memo " + memo )
       thisstudent = student_dictionary[ studentno ] 

       # Now copy the memo to my stash of memos and rename to student number
       shutil.copy( '/Users/gareth/Desktop/DS/Downloadedmemos/' + memo.replace(".docx",".pdf"), "Memos/" + studentno + ".pdf" )
       # Copy the students memo for the advisor of study 
       if "advisor" in thisstudent.keys() : copyMemo( studentno, thisstudent["advisor"], "advisor", staffdict[thisstudent["advisor"]] )
       # Copy the students memo for the personal tutor
       if "personal" in thisstudent.keys() : copyMemo( studentno, thisstudent["personal"], "personal", staffdict[thisstudent["personal"]] )
       # Copy the students memo for the supervisor
       if "supervisor" in thisstudent.keys() : 
           print("CHECKING ",  thisstudent )
           copyMemo( studentno, thisstudent["supervisor"], "supervisor", staffdict[thisstudent["supervisor"]] )
       # Now make copies for modules
       for module in thisstudent["modules"] :
           for staff in modteach[module] : 
               copyMemo( studentno, staff, module, staffdict[staff] )
       print( "Made memo for student " + studentno + " with memo " + memo )
