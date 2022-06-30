# SecTool
I developed Security Tool for Windows with python in order to analyise system more efficiently. I used VirusTotal API, Windows System Internals.
<br>
|<img src="Images/AnaEkran.png">|
|---------|
| Main Menu |
<br>

## To Run Project in Terminal
- ### The API Key for VirusTotal is mine. You can replace it with yours if you have advantaged profile in VirusTotal.
- ### If you want to run the scipt in terminal, you have to look at <a href="https://github.com/mrtergl/VirusTotal_Python_URL_Scan">this link</a>.
- ### Do not change the folder and file names in python script.
- ### All files will be at C:/Users/{user}/SecTool.

<br>

## Necessary Libraries:

| Library  | Code to install |
| ------------- | ------------- |
| hashlib  | ``` pip install hashlib ```  |
| pathlib2  | ``` pip install pathlib2 ```  |
| WMI  |``` pip install WMI ``` |
| requests | ``` pip install requests ```|
| browser-history 0.3.2 | ``` pip install browser-history ```|
| psutil 5.9.1 | ``` pip install psutil ```|
| python-cfonts 1.5.2 | ``` pip install python-cfonts ```|
| colorama 0.4.5 | ``` pip install colorama ```|
| pywin32 304 | ``` pip install pywin32 ```|
| openpyxl 3.0.10 | ``` pip install openpyxl ```|
<br>

## Scan Browser History with VirusTotal

#### In this option you can scan all the browser history in device using Virus Total. This is the updated version of one of my repos <a href="https://github.com/mrtergl/VirusTotal_Python_URL_Scan">URL_SCAN</a>.

|<img src="Images/1-a.png">|<img src="Images/1-b.png">|
|---------|---------|
| Terminal view | Code |

<br>
There will be "history.csv" file for all the browser histories.

#### After scanning. You can see the results in C:/Users/{username}/SecTool/vt_results.txt .

|<img src="Images/1-c.png">|
|---------|
| Results |

<br>

## Scan Running Processes with their hash in VirusTotal

#### In this option you can scan all the processes hash running in device using Virus Total. This is the updated version of one of my repos <a href="https://github.com/mrtergl/VirusTotal_Process_Scan">Process_Scan</a>.

|<img src="Images/2-a.png">|<img src="Images/2-b.png">|
|---------|---------|
| Terminal view | Code |

It takes a bit long because it searchs for the absolute path of the process to get the hash of it. After finding the hash, it scans in the Virustotal. You can find the scanned process in "vt_Results_exe.txt". After all the process scanned. It will be shown at an excel file called "vt_Results_exe.xlsx".

|<img src="Images/2-c.png">|<img src="Images/2-d.png">|
|---------|---------|
| Text File | Excel File |

#### Note: .txt file will be converted to .xlsx file after scanning.

## Scan all the foreign IP addresses that communicating with device in VirusTotal

## Get Netstat Connection Table and Foreign IP addresses
## Get all the Windows Services
## Get Startup Files

