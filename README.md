# SSWSysAdmins.LoginScript
## SysAdmin - This login script does:

   1. Flushes DNS
   2. Copies Office Templates to your machine, as per the rule https://rules.ssw.com.au/have-a-companywide-word-template
   3. If you have Snagit installed, copies Snagit Template to your machine, then opens the SSW.snagtheme so Snagit registers the SSW theme! As per the rule https://www.ssw.com.au/rules/screenshots-add-branding'

## How to use it:

1. Close any Office applications open (or else the templates will not be copied);
2. Start a PowerShell session
3. Run:
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex (new-object net.webclient).downloadstring('https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/main/Script/SSWLoginScript.ps1')
```
4. The script should copy everything, close itself and open a notepad with the log when it's done.
	
Note: Some red errors may appear. If you have any problems, ask Kiki at kiki@ssw.com.au.
	

This script is designed to be run from PowerShell directly, with the link above. Let's see why:

   ✅ Same command for domain-joined and BYOD machines 
   
   ✅ Open source - everyone can improve it
   
   ✅ Not file server dependent - doesn't require VPN and can be run from anywhere
   
   ❌ Manual - It's not run automatically, the above needs manual action  

**If you have reset your PC, you need to remember to re-run the script!**

![Reset PC](/Images/ResetPC1.png)

**Figure: If you clicked on “Reset this PC”, you need to re-run the script**


## Owner: [Kaique Biancatti](https://www.ssw.com.au/people/kaique-biancatti)
