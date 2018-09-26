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


# Open the worksheet that contains all our useful information on the students
wbp = xw.Book( '/Users/gareth/Desktop/DS/2018/basic-information.xlsm')
# Get a sheet from the workbook
sht = xw.sheets['options']
# Get data on staff numbers from exccel and make a dictionary
n, staffdict, staffdata, staffnumbers = 0, {}, sht.range("G3:G72").value, sht.range("F3:F72").value
for name in staffdata : 
    staffdict[ name ] = str( int(staffnumbers[n]) )
    n += 1

# for k,n in staffdict.items() : print( k, n )

# List of possible student requirements
possible_requirements = ["Green Room", "Individual room", "Cubical 6-8", "font size 18 on A4 paper"]

# Make a dictionary of who teaches each module
n, modteach, modules, teachers = 0, {}, sht.range("K3:K82").value, sht.range("L3:L82").value
for module in modules : 
    mod = module.replace("*","")
    modteach[mod] = {}
    if module.find("*")!=-1 : modteach[mod]["project"] = True
    else : modteach[mod]["project"] = False
    modteach[mod]["project_students"] = [] 
    modteach[mod]["teachers"] = teachers[n].split()
    for requirement in possible_requirements : modteach[mod][requirement] = []
    # Check we have staff numbers for everyone who teaches a module
    for staff in modteach[mod]["teachers"] :
        if staff not in staffdict : RuntimeError("staff number not found for " + staff)
    n+=1 

# for k, n in modteach.items() : print( k, n )

# Get the list of advisors of study
advisor_names, advisor_dictionary = sht.range("D3:D12").value, {}
for advisor in advisor_names :
    # Open advisors excel sheet
    shd =  xw.sheets[advisor]
    # Read data on students
    n, advisor_dictionary[advisor], s_data = 0, {}, shd.range('B8').expand().value
    for col in s_data[0] :
        if col == "fake" : continue
        tdict = {}
        tdict["modules"] = []
        for mod in range(1,7) :
            if (s_data[mod][n]!=None) & (s_data[mod][n]!="End") & (s_data[mod][n]!="end") : tdict["modules"].append( s_data[mod][n] )
        advisor_dictionary[advisor][ str(int(col)) ] = tdict 
        n += 1

# Read in the main sheet
shd = xw.sheets["main"]
n, student_dictionary, data = 0, {}, shd.range('B8').expand().value
for col in data[0] :
    tdict, advisor = {}, ""
    tdict["supervisor"], tdict["modules"] = [], []
    if (data[4][n]!=None) & (data[4][n]!="End") : tdict["advisor"], advisor = data[4][n], data[4][n] 
    if (data[5][n]!=None) & (data[5][n]!="End") : tdict["personal"] = data[5][n]
    if (data[6][n]!=None) & (data[6][n]!="End") : tdict["supervisor"].append( data[6][n] )
    if (data[7][n]!=None) & (data[7][n]!="End") : tdict["supervisor"].append( data[7][n] )
    if (data[8][n]!=None) & (data[8][n]!="End") : tdict["supervisor"].append( data[8][n] )
    # Now get modules from advisor dictionary
    if advisor != "" :
       for module in advisor_dictionary[advisor][ str(int(col)) ]["modules"] : tdict["modules"].append( module )
    student_dictionary[ str(int(col)) ] = tdict 
    n += 1

#print( student_dictionary ) 

# # Now work through all the memos that were downloaded
os.mkdir("Memos")
os.mkdir("Allstaff")
for memo in os.listdir("/Users/gareth/Desktop/DS/2018/Downloadedmemos/") :
    if memo.endswith(".docx") :
       # Get the student number from the memo
       text = textract.process('/Users/gareth/Desktop/DS/2018/Downloadedmemos/' + memo ).decode("utf-8")
       start = text.find("Student No:")
       end = text.find("Course:")
       studentno = text[start:end].replace("Student No:","").rstrip().strip()

       # Now find the details of the student from excel
       print("Making memos for student number " + studentno + " with memo " + memo )
       if studentno not in student_dictionary : RuntimeError("Could not find student number " + studentno + " for memo " + memo )
       thisstudent = student_dictionary[ studentno ] 

       # Now copy the memo to my stash of memos and rename to student number
       shutil.copy( '/Users/gareth/Desktop/DS/2018/Downloadedmemos/' + memo.replace(".docx",".pdf"), "Memos/" + studentno + ".pdf" )
       # Copy the students memo for the advisor of study 
       if "advisor" in thisstudent.keys() : copyMemo( studentno, thisstudent["advisor"], "advisor", staffdict[thisstudent["advisor"]] )
       # Copy the students memo for the personal tutor
       if "personal" in thisstudent.keys() : copyMemo( studentno, thisstudent["personal"], "personal", staffdict[thisstudent["personal"]] )
       # Copy the students memo for the supervisor
       for tsuper in thisstudent["supervisor"] : copyMemo( studentno, tsuper, "supervisor", staffdict[tsuper] )
       # Now make copies for modules
       for module in thisstudent["modules"] :
           # Nothing to do for computer science modules 
           if "CSC" in module : continue
           # Check all things we need to note down
           for requirement in possible_requirements :
               if( text.find(requirement)!=-1 ) : modteach[module][requirement].append( studentno )
           # Check if this is a project module and add to list of project students if it is
           if modteach[module]["project"] : modteach[module]["project_students"].append( studentno )
           # Coopy the modules we need
           for staff in modteach[module]["teachers"] : 
               copyMemo( studentno, staff, module, staffdict[staff] )
       print( "Made memo for student " + studentno + " with memo " + memo )

# Get information for teaching office on green room student in each module
os.mkdir("Office")
for module, dicto in modteach.items() :
    # Get project students
    if dicto["project"] : 
       print("STUDENTS TAKING PROJECT MODULE " + module )
       for student in dicto["project_students"] : print( student )
    # Print info for teaching office
    of = open('/Users/gareth/Desktop/DS/2018/Office/' + module + ".txt", "w")
    for requirement in possible_requirements :
        if len(dicto[requirement])>0 : 
           of.write("Requirement: " + requirement + "\n" )
           of.write("--------------------------------- \n")
           of.write("\n")
           for student in dicto[requirement] : of.write( student + " ")
           of.write("\n")
           of.write("\n")
    of.close()
