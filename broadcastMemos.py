import requests
import os
from getpass import getpass
from requests_ntlm import HttpNtlmAuth

# TODO : This script needs to be reworked so that it deletes from all directories on sharepoint first and then uploads from files on my computer

api_url = "https://intranet.qol.qub.ac.uk/schools/smp/education/_api"

sp_user = input('User: ')
sp_password = getpass('Password: ')

session = requests.Session()
session.auth = HttpNtlmAuth(sp_user, sp_password)
session.headers.update({"Accept": "application/json; odata=verbose"})

with session as s:
    for target_dir in os.listdir("Allstaff/") :
         if target_dir==".DS_Store" : continue
         # Check if the directory exits
         r = s.get( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/" + target_dir + "')")
         if r.status_code!=200 : raise Exception('failed to find directory ' + target_dir + " on sharepoint")         

         # Get a list of files in the folder.
         r = s.get( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/" + target_dir + "')/Files" )
         print("In " + target_dir + " found files : ") 
         for sp_file in r.json()['d']['results'] : print(sp_file['Name'])
         dd = input('Do you want to delete these files from ' + target_dir + ": ")

         # Get a form digest value.
         # See https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
         r2 = s.post( api_url + "/contextinfo")
         form_digest = r2.json()['d']['GetContextWebInformation']['FormDigestValue']

         if dd=='y' : 
            for sp_file in r.json()['d']['results'] : 
                print( "Deleting " + target_dir + "/" + sp_file['Name'] )
                r3=s.post( api_url + "/web/GetFileByServerRelativeUrl('/schools/smp/education/Disability/Memos/" + target_dir + "/" + sp_file['Name'] + "')", headers = {"X-RequestDigest": form_digest, "X-HTTP-Method":"DELETE"} )
                if r3.status_code!=200 : raise Exception('failed to delete ' + sp_file['Name'] + " on sharepint")

         # Post all the files
         for filename in os.listdir("Allstaff/" + target_dir) :
             print( "Uploading " + target_dir + "/" + filename + " to sharepoint")
             with open("Allstaff/" + target_dir + "/" + filename, 'rb') as f :
                r3 = s.post( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/" + target_dir + "')/Files/add(url='" + filename.replace("*","") + "')", headers = {"X-RequestDigest": form_digest},  files = {"file": f}) 
                if r3.status_code!=200 : raise Exception('failed to upload ' + filename + " on sharepint")

    # Check if teaching office directory exits
    r = s.get( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/Teaching_office')")
    if r.status_code!=200 : raise Exception('failed to find Teaching_office directory on sharepoint')

    # Get a list of files in the folder.
    r = s.get( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/Teaching_office')/Files" )
    print("In Teaching_office found files : ")
    for sp_file in r.json()['d']['results'] : print(sp_file['Name'])
    dd = input('Do you want to delete these files from Teaching_office : ')

    r2 = s.post( api_url + "/contextinfo")
    form_digest = r2.json()['d']['GetContextWebInformation']['FormDigestValue']
    
    if dd=='y' : 
       for sp_file in r.json()['d']['results'] :
           print( "Deleting Teaching_office/" + sp_file['Name'] )
           r3=s.post( api_url + "/web/GetFileByServerRelativeUrl('/schools/smp/education/Disability/Memos/Teaching_office/" + sp_file['Name'] + "')", headers = {"X-RequestDigest": form_digest, "X-HTTP-Method":"DELETE"} )
           if r3.status_code!=200 : raise Exception('failed to delete ' + sp_file['Name'] + " on sharepint")
    
    # Post all the files
    for filename in os.listdir("Office") : 
        print( "Uploading Office/" + filename + " to sharepoint")
        with open("Office/" + filename, 'rb') as f :
           r3 = s.post( api_url + "/web/GetFolderByServerRelativeUrl('Disability/Memos/Teaching_office')/Files/add(url='" + filename.replace("*","") + "')", headers = {"X-RequestDigest": form_digest},  files = {"file": f})
           if r3.status_code!=200 : raise Exception('failed to upload ' + filename + " on sharepint")


