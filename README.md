# Win_AUMID_dict
Script that finds the Application User Model ID or AUMID of an installed app in Windows 10 and creates a dictionary via Excel with name of application as key and AUMID as value

Note:
Before using this, you should increase character limit per line in Windows Powershell, otherwise most AUMIDs will be incomplete. Here's how to do it:

Open Windows PowerShell.
Right click the top bar > Properties > Layout > Screen Buffer Size > Wrap text output on resize: uncheck + increase Width value, e.g. to 1200

<img width="460" alt="powershell width" src="https://user-images.githubusercontent.com/25702508/206517557-607f8f9b-b559-4271-b77f-f0bd7764c7d6.PNG">


Works kind of ok, but mixes up some values at the moment.
