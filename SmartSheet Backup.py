import smartsheet
import json
import zipfile
import shutil
import time
import os

token = TOKEN SYSADMIN
ss = smartsheet.Smartsheet(token)
backupdir = "d:\\Backup"
if not os.path.exists(backupdir):
    os.mkdir(backupdir)
zipname = "SmartsheetBackup - " + time.strftime('%Y-%m-%d-%H-%M')
odir = "d:\\Backup\\SmartsheetBackup - " + time.strftime('%Y-%m-%d-%H-%M')
rapportname = zipname + " Report.txt"
skip = 0
users = ss.Users.list_users(include_all=True).data
ss.errors_as_exceptions(True)

##Création Dossier de backup, rapport de backup

os.mkdir(odir)
rapport = open(odir + "\\" + rapportname, "w+")
rapport.write("---------- Starting Backup ----------\n\n")
rapport.write(time.strftime('%Y-%m-%d-%H-%M\n\n'))

def Zipall():
    
    #Initialisation Zip Dossier Backup + Fichier rapport

    os.chdir(backupdir)
    print("Zipping Backup")
    rapport.write("Zipping Archive")
    rapport.write("\n\nRapport end at : " + time.strftime('%Y-%m-%d-%H-%M') + "\n")
    rapport.write("----------- Ending Backup -----------")
    rapport.write(str(skip) + " Users Skipped")
    rapport.close()
    shutil.make_archive(zipname, format="zip", root_dir=odir)
    print("Zipping Complete !")

    #Fin de Zipping Dossier

    #Néttoyage fichier et dossier

    shutil.rmtree(odir)
    
def BackupTool(currentuser):

    ssh = ss.Home.list_all_contents()
    lsthome = json.loads(ssh.to_json())
    sss = ss.Sheets.list_sheets(include_all=True) 
    lstsheetsall = json.loads(sss.to_json())
    lstsheets = []
    workspaces = ss.Workspaces.list_workspaces(include_all=True)
    lstworkspaces = json.loads(workspaces.to_json())
    
    #Backup des dossiers et sheets Inférieurs ou égales à F+2
    
    rapport.write("Backup " + currentuser)
    print("Backup en cours : " + currentuser)
    os.mkdir(odir + "\\" + currentuser)
    os.chdir(odir + "\\" + currentuser)
    print("Tool About to backup folder and sheets equals and less than f+2")
    rapport.write("Start Backup Folders <= F+2\n\n")
    for folder in lsthome['folders']:
        rapport.write("Create Folder : " + folder['name'] + "\n")
        os.mkdir(odir + "\\" + currentuser + "\\" + folder['name'])
        os.chdir(odir + "\\" + currentuser + "\\" + folder['name'])
        cdir = os.getcwd()
        for sheets in folder['sheets']:
            try:
                rapport.write("Download : " + sheets['name'] + "\n")
            except:
                rapport.write("Erreur : " + "\n")
            ss.Sheets.get_sheet_as_excel(sheets['id'], os.getcwd())
            lstsheets.append(sheets['id'])
        if folder['folders'] == '[]':
            break
        else:
            for sfolder in folder['folders']:
                rapport.write("Create Folder : " + sfolder['name'] + "\n")
                os.mkdir(odir + "\\" + currentuser + "\\" + sfolder['name'])
                os.chdir(odir + "\\" + currentuser + "\\" + sfolder['name'])
                for ssheets in sfolder['sheets']:
                    rapport.write("Download : " + ssheets['name'] + "\n")
                    ss.Sheets.get_sheet_as_excel(ssheets['id'], os.getcwd())
                    lstsheets.append(ssheets['id'])
    rapport.write("Folder and sheets f+2 Backup OK \n\n")
    print("Folders And Sheets Backup OK !")

    #Fin De Backup des Dossiers

    #Initialisation du Backup des fichiers root

    print("Tool About to backup Root Sheets")
    rapport.write("Download Root Sheets \n")
    for bsheets in lsthome['sheets']:
        os.chdir(odir + "\\" + currentuser)
        ss.Sheets.get_sheet_as_excel(bsheets['id'], os.getcwd())
        rapport.write("Download Rsheets : " + bsheets['name'] + "\n")
        lstsheets.append(bsheets['id'])
    rapport.write("Root Sheets Backup OK \n\n")
    print("Root Sheets Bakcup OK !")

    #Fin de backup des fichier root

    #Backup des sheets restants dans SheetsSup

    os.chdir(odir)
    print("Tool About to backup file in more than f+2")
    os.mkdir(odir + "\\" + currentuser + "\\SheetsSup")
    os.chdir(odir + "\\" + currentuser + "\\SheetsSup")
    rapport.write("\n\nBackup sheets over f+2\n\n")
    if not(str(lstsheetsall['totalCount']) == str(len(lstsheets))):
        for sheetsall in lstsheetsall['data']:
            if not(sheetsall['id'] in lstsheets):
                try:
                    rapport.write("Download : " + sheetsall['name'] + "\n")
                except UnicodeEncodeError:
                    rapport.write("Erreur : " + "\n")
                try:
                    ss.Sheets.get_sheet_as_excel(sheetsall['id'], os.getcwd())
                except ConnectionResetError:
                    print("Error Conn")
                    continue
        os.chdir(odir + "\\" + currentuser)
    rapport.write("Sheets Backup OK \n")
    print("Sheets Backup Sup OK !")

    #Fin backup Sheets over F+2

    #Initialisation Backup Attachments

    rapport.write("\n\nStart Backup Attachments\n\n")
    print("Tool About to backup Attachments")
    os.mkdir(odir + "\\" + currentuser + "\\Attachments")
    os.chdir(odir + "\\" + currentuser + "\\Attachments")
    recu = 0
    for sheetsall in lstsheetsall['data']:
        t = True
        try:
            DLA = ss.Attachments.list_all_attachments(sheetsall['id'], include_all=True)
        except ValueError:
            print("Cannot Backup Attachs")
            rapport.write("Skip attachs")
            t = False
        except smartsheet.exceptions.UnexpectedRequestError:
            print("Cannot Backup Attachs")
            rapport.write("Skip attachs")
            t = False
        if(t == True):
            action = DLA.data
            altn = ""
            for col in action:
                rapport.write("Download : " + col.name + "\n")
                attachurl = ss.Attachments.get_attachment(sheetsall['id'],col.id)
                path = os.getcwd() + "\\" + attachurl.name
                if not os.path.exists(path):
                    try:
                        ss.Attachments.download_attachment(attachurl, os.getcwd())
                        break
                    except:
                        print("Can't Download the Attachment")
                else:
                    recu += 1
                    altn = "(" + str(recu) + ")" + col.name
                    try:
                        ss.Attachments.download_attachment(attachurl, os.getcwd(), alternate_file_name=altn)
                        break
                    except:
                        print("Can't dl this file")
    print("Download Attachments OK !")
    rapport.write("Download Attach OK")
    os.chdir(odir)

    #Fin de backup attachments

for i in users:
    x = True
    z = True
    try:
        ss.assume_user(i.email)
    except:
        print("probleme assume")
        z = False
    if(i.name == None):
        try:
            
            BackupTool(i.email)
            
        except ConnectionResetError:
            
            print("Error Connection Aborted")
            time.sleep(3)
            shutil.rmtree(odir + "\\" + currentuser)
            print("Retry")
            time.sleep(10)
            
            try:
                
                Backuptool(i.email)
                
            except ConnectionResetError:
                
                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break
            
            except smartsheet.exceptions.UnexpectedRequestError:

                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break

            except KeyboardInterrupt:

                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break
            
        except smartsheet.exceptions.UnexpectedRequestError:
            
            print("Error Connection Aborted")
            time.sleep(3)
            shutil.rmtree(odir + "\\" + currentuser)
            print("Retry")
            time.sleep(10)
            
            try:
                
                Backuptool(i.email)
                
            except ConnectionResetError:
                
                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break
            
            except smartsheet.exceptions.UnexpectedRequestError:

                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break

            except KeyboardInterrupt:

                ol = False
                print("Error Connection Aborted")
                print("User " + i.email + " Skipped")
                break
            
        except KeyboardInterrupt:
            
            ol = False
            print("Error Connection Aborted")
            Zipall()
            exit(True)
            
    else:
        
        if(z == True):
            
            while(x == True):
                
                try:
                    
                    ss.Home.list_all_contents()
                    y = True
                    break
                
                except:
                    
                    try:
                        
                        print("User " + i.name + " cannot backup")
                        
                    except:
                        
                        print("cannot backup user")
                    x = False
                    y = False
                    
            if(y == True):
                
                ol = True
                
                try:
            
                    BackupTool(i.name)
            
                except ConnectionResetError:
                    
                    print("Error Connection Aborted")
                    time.sleep(3)
                    shutil.rmtree(odir + "\\" + currentuser)
                    print("Retry")
                    time.sleep(10)
                    
                    try:
                        
                        Backuptool(i.name)
                        
                    except ConnectionResetError:
                        
                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break
                    
                    except smartsheet.exceptions.UnexpectedRequestError:

                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break

                    except KeyboardInterrupt:

                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break
                    
                except smartsheet.exceptions.UnexpectedRequestError:
                    
                    print("Error Connection Aborted")
                    time.sleep(3)
                    shutil.rmtree(odir + "\\" + currentuser)
                    print("Retry")
                    time.sleep(10)
                    
                    try:
                        
                        Backuptool(i.name)
                        
                    except ConnectionResetError:
                        
                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break
                    
                    except smartsheet.exceptions.UnexpectedRequestError:

                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break

                    except KeyboardInterrupt:

                        print("Error Connection Aborted")
                        print("User " + i.name + " Skipped")
                        skip += 1
                        break
                    
                except KeyboardInterrupt:
                    
                    print("Error Connection Aborted")
                    Zipall()
                    exit(True)
                    
            elif(y == False):
                
                skip += 1
                print("not backup")
                
if(ol == True):
    
    Zipall()
    
exit(True)

#Fin de Programme
