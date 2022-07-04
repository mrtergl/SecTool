from ast import Try
from genericpath import exists
import hashlib
from browser_history import get_history
import requests
import json
import base64
import time
import wmi
from pathlib import Path
import pandas as pd
from re import T
import re
import psutil
from pprint import pprint as pp
from cfonts import render
from subprocess import check_output
import os
import win32com.client as win32
from colorama import Fore, Back, Style



home = str(Path.home())

try: 
    os.mkdir(home+"/SecTool") 
except OSError as error: 
    pass

home = home+"\SecTool"

headers = {
        "Accept": "application/json",
        "x-apikey": "# YOUR OWN API KEY !!!" 
    }

x = True

def TxtToExcel(file):

    df = pd.read_csv(home+file+".txt", sep=',')
    os.remove(home+file+".txt")
    if exists(home+file+".xlsx"):
        os.remove(home+file+".xlsx") 
    df.to_excel(home+file+'.xlsx', index=False)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(home+file+".xlsx")
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()



def IP_STATS():

    open(home+"/Netstat.txt", "w").close()
    
    def Get_Ip():
        IP_List = set()
        with open(home+"/Netstat.txt", "r") as f_in:
            for line in map(str.strip, f_in):
                if not line or line == "State":
                    continue
                line = re.split(r"\s{2,}", line)
                if (len(line) == 5 and not "Foreign Address" in line[2] and not "0.0.0.0"
                    in line[2] and not "127.0.0.1" in line[2] and not "[::]" in line[2]):
                    if ":" in line[2]:
                        IP_List.add(line[2][:line[2].find(":")])
                    else:
                        IP_List.add(line[2])
                else:
                    continue
        return IP_List

    url = "https://www.virustotal.com/api/v3/ip_addresses/"

    out = check_output(["netstat", "-ano"])
    x = str(out, "utf-8")
    with open(home+"/Netstat.txt",'a') as vt:
        vt.write(x)
    
    os.system("start "+home+"/Netstat.txt")
    #
    ans = input(str(Style.BRIGHT  + Fore.GREEN +"\nDo you want scan all the foreign IP Adresses in Virus Total ?[y/n]: "+Style.RESET_ALL))

    if ans == "y":
        with open(home+"/IP_RESULTS.txt",'a') as vt:
            vt.write ('IP,Harmless,Malicious,Suspicious,Undetected,Timeout\n\n')
        Ip_List_set= Get_Ip()
        Ip_List = list(Ip_List_set)
        Length = len(Ip_List)
        for ip in range(Length):
            url = "https://www.virustotal.com/api/v3/ip_addresses/"
            newurl = url+Ip_List[ip]
            response = requests.get(newurl, headers=headers)
            x = json.loads(response.content)
            if len(x["data"]) != 0:
                data =  x["data"]["attributes"]["last_analysis_stats"]

                with open(home+"/IP_RESULTS.txt",'a') as vt:
                    vt.write("{}".format(Ip_List[ip]))
                    for value in data.values():
                        vt.write(',{}'.format(value))
                    vt.write("\n")
                print("[{}/{}] DONE".format(ip+1,Length))
            else:
                continue
        
        TxtToExcel("/IP_RESULTS")

        ans = input(str(Style.BRIGHT  + Fore.GREEN +"\nProcess has done. Would you like to open scanned IPs and results ? [y/n]: "+Style.RESET_ALL))
        
        if ans == "y":
            os.system("start "+home+"/IP_RESULTS.xlsx")
            menu()
        else:
            menu()
        
    elif ans == "n":
        menu()



def Browser_History():

    print("\n")
    outputs = get_history()
    his = outputs.histories
    outputs.save(home+"\history.csv")


    current_directory = os.getcwd()

    liste =[]
    col_list = ["URL", "Timestamp"]
    df = pd.read_csv(home+'\history.csv', usecols=col_list)
    for i in range(len(df.index)):
        
        liste.append(df["URL"].iloc[i])
        
    i=0

    for site in liste:
        url = "https://www.virustotal.com/api/v3/urls/"
        url_id = base64.urlsafe_b64encode(site.encode()).decode().strip("=")
        url = url+url_id
        response = requests.get(url, headers=headers)
        c = response.text
        x = json.loads(response.content)
        if "data" in x:
            data = x["data"]["attributes"]["last_analysis_stats"]["malicious"]
            if data <= 0:
                with open(home+'/vt_results.txt','a') as vt:
                    vt.write(site) and vt.write (' -\tNOT MALICIOUS\n')
                    print("[{}/{}] DONE".format(i+1,len(liste[i])))
            elif 1 <= data >= 3:
                with open(home+'/vt_results.txt','a') as vt:
                    vt.write(site) and vt.write (' -\tMAYBE MALICIOUS\n')
                    print("[{}/{}] DONE".format(i+1,len(liste[i])))
            elif data >= 4:
                with open(home+'/vt_results.txt','a') as vt:
                    vt.write(site) and vt.write (' -\tMALICIOUS\n')
                    print("[{}/{}] DONE".format(i+1,len(liste[i])))
            else:
                print("url not found")
        else:
            print("url not found")
            
        i = i+1    
        time.sleep(15)

    print("SCANNING COMPLETED")



def HASH_SCAN():

    print("\n")
    liste = []
    f = wmi.WMI()
    for process in f.Win32_Process():
        file_extension = Path(process.Name).suffix
        if file_extension == '.exe':
            liste.append(process.Name)
        
    BUF_SIZE = 32768
    md5 = hashlib.md5()
    sha1 = hashlib.sha1()
    def find(name):
        for root, dirs, files in os.walk("C:\\"):
            if name in files:
                return os.path.join(root, name)

    
    with open(home+"/vt_Result_exe.txt",'a') as vt:
        vt.write ('Name,Hash,Harmless,Type Unsupported,Suspicious,Confirmed Timeout,Confirmed Timeout,Failure,Malicious,Undetected\n\n')


    for i in range(len(liste)):
        path=find(liste[i])
        if type(path) == str:
            try:
                with open(path,"rb") as f:
                    bytes = f.read()
                    readable_hash = hashlib.md5(bytes).hexdigest()
                url = "https://www.virustotal.com/api/v3/search?query="
                url = url+readable_hash
                response = requests.get(url, headers=headers)
                x = json.loads(response.content)
                del x["links"]
                if len(x["data"]) != 0:
                    data =  x["data"][0]["attributes"]["last_analysis_stats"]
                    with open(home+"/vt_Result_exe.txt",'a') as vt:
                        vt.write("{},{}".format(liste[i],readable_hash))
                        for value in data.values():
                            vt.write(',{}'.format(value))
                        vt.write("\n")
                    print("[{}/{}] DONE".format(i+1,len(liste)))
                else:
                    with open(home+"/vt_Result_exe.txt",'a') as vt:
                        vt.write("{},Built-in-Service".format(liste[i]))
                        vt.write("\n")
                    print("[{}/{}] DONE".format(i,len(liste)))
            except OSError:
                with open(home+"/vt_Result_exe.txt",'a') as vt:
                    vt.write("{},{},DOSYA YOLU OKUNAMADI".format(liste[i],readable_hash))
        else:
            with open(home+"/vt_Result_exe.txt",'a') as vt:
                vt.write("{},NOT FOUND".format(liste[i]))
                vt.write("\n")
            print("[{}/{}] DONE".format(i,len(liste)))
           
    TxtToExcel("/vt_Result_exe")
        
    print("SCANNING COMPLETED")



def win32_service():

    def getService():
        with open(home+"/win32_services.txt",'w') as vt:
            vt.write('Name,Display Name,Process ID, Status\n\n') 
        services = psutil.win_service_iter()
        with open(home+"/win32_services.txt",'a',encoding="utf-8") as vt:
            for x in services:
                vt.write("{},{},{},{}\n".format(x.name(),x.display_name(),x.pid(),x.status()))

    getService()

    TxtToExcel("/win32_services")

    ans = input(str(Style.BRIGHT  + Fore.GREEN +"\nService file created.Would you like to open the file to see all services? [y/n]: "+Style.RESET_ALL))

    if ans == "y":
            os.system("start "+home+"\win32_services.xlsx")
            menu()
    else:
        menu()


def Startups():

    print(Style.BRIGHT+Fore.LIGHTYELLOW_EX+"\nLoading...\n"+Style.RESET_ALL)

    c= wmi.WMI()

    wql = "SELECT * FROM Win32_StartupCommand"

    with open(home+"/Startups.txt",'a') as vt:
        vt.write("Name,Caption,Description,User,Location\n")
        for x in c.query(wql):
            vt.write("{},{},{},{},{}".format(x.Name,x.Caption,x.Description,x.User,x.Location))
            vt.write("\n")
    
    TxtToExcel("/Startups")

    print(Style.BRIGHT+Fore.LIGHTYELLOW_EX+"DONE"+Style.RESET_ALL)
    ans = input(str(Style.BRIGHT  + Fore.GREEN +"\nStartup file created. Would you like to open the file ? [y/n]: "+Style.RESET_ALL))

    if ans == "y":
            os.system("start "+home+"\Startups.xlsx")
            menu()
    else:
        menu()

    


def number_to_string(argument):
    match argument:
        case 0:
            exit()
        case 1:
            return Browser_History()
        case 2:
            return HASH_SCAN()
        case 3:
            return IP_STATS()
        case 4:
            return win32_service()
        case 5:
            return Startups()



def menu():
    while(True):
        output = render('SecTool', colors=['green', 'white'], align='center', font='slick')
        print(output)
        print (Style.BRIGHT+Fore.YELLOW+'1'+Style.RESET_ALL+' -- VirusTotal Browser History Control'+Style.RESET_ALL)
        print (Style.BRIGHT+Fore.YELLOW+'2'+Style.RESET_ALL+' -- VirusTotal Process Control'+Style.RESET_ALL )
        print (Style.BRIGHT+Fore.YELLOW+'3'+Style.RESET_ALL+' -- Get Netstat Connection Table and Foreign IP addresses'+Style.RESET_ALL )
        print (Style.BRIGHT+Fore.YELLOW+'4'+Style.RESET_ALL+' -- Get all the Windows Services'+Style.RESET_ALL )
        print (Style.BRIGHT+Fore.YELLOW+'5'+Style.RESET_ALL+' -- Get Startup Files'+Style.RESET_ALL )
        print (Style.BRIGHT+Fore.YELLOW+'0'+Style.RESET_ALL+' -- Exit\n' )
        option = int(input(Style.BRIGHT+Fore.LIGHTGREEN_EX+'Enter Your Choice: '+Style.RESET_ALL))
        number_to_string(option)

        

menu()



       
