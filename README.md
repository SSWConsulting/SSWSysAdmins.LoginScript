# LoginScript
SysAdmin - This login script pulls default Office and SnagIt templates from GitHub into employee's computers.

How to use it:

	1. Close any Office applications open (or else the templates will not be copied);
	2. Start a PowerShell session as an Administrator
	3. Run:
```powershell
iex (new-object net.webclient).downloadstring('https://github.com/SSWConsulting/LoginScript/raw/master/Script/SSWLoginScript.ps1')
```
	4. Click 'Ok' to the prompts, if any;
	5. If you are not logged in as a SSW user, input your username on the pop-up that appears;
	6. The script should copy everything, close itself and open a notepad with the log when it's done.
	
	Note: Some red errors may appear. If you have any problems, ask Kaique at KaiqueBiancatti@ssw.com.au.
