# LoginScript
SysAdmin - This login script does:

   1. Flushes DNS
   2. Synchronizes your PC time with the computer time of the Sydney server
   3. Copies Office Templates to your machine, as per the rule https://rules.ssw.com.au/have-a-companywide-word-template
     a. If you do not have access to our fileserver, copies them from GitHub
   4. Copies Outlook Signatures to your PC (using the same rules as above)
   5. Closes SnagIt if it was open, and copied its templates to your PC (using the same rules as above)
   6. Changes the desktop background image to be SSW, if user wants to do so

How to use it:

	1. Close any Office applications open (or else the templates will not be copied);
	2. Start a PowerShell session as an Administrator
	3. Run:
```powershell
iex (new-object net.webclient).downloadstring('https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/main/Script/SSWLoginScript.ps1')
```
	4. Click 'Ok' to the prompts, if any;
	5. If you are not logged in as a SSW user, input your username on the pop-up that appears;
	6. The script should copy everything, close itself and open a notepad with the log when it's done.
	
	Note: Some red errors may appear. If you have any problems, ask Kaique at KaiqueBiancatti@ssw.com.au.

Owner: [Kaique Biancatti](https://www.ssw.com.au/people/kaique-biancatti)
