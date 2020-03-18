# Script-Manager
For use with the Google Sheets Add-on named Script Manager to manage Apps Script files


Contact Info	
Email:	aj.addons@gmail.com
	
Set Up Instructions	
	
You will need to create a bound script with provided code, authorize the script, create a GCP project, enable an API, enable an Advanced Service, and associate the GDP project with the Apps Script project.  I'm providing the code and instructions for how to do the set-up.
	
Create an Apps Script project bound to the spreadsheet that the add-on is installed in

	From the spreadsheet menu, click the "Tools" menu and the "Script editor" menu item.
	The code editor will open
	In the upper left hand corner, click the name - Enter a Name: Script_Manager_Bound
	Go to the code source at the link below
	
	Copy the code in the Script_Manager_User.gs file
	Paste the code into the code editor
	Save the File
	
	From the code editor, click the View menu and click "Show manifest file"
	The appsscript.json file will open - you will see it displayed in the code editor
	Copy the appsscript.json file from GitHub and paste it into your appsscript.json file
	Save the file
	
	Got to Link:
	https://console.cloud.google.com/home/dashboard?authuser=0
	Click the drop down list of project names at the top
	In the upper right hand corner click New Project
	Enter a project name
	If you are in a GSuite account and don't have an organization name then create an organization name
	Click Create
	Navigate to APIs and Services
	Click Enable APIs and Services button at top of page
	Search: Apps Script API
	Click: Apps Script API
	Click: Enable
	Get the new GCP project number
	Go back to the code editor
	Click the Resource menu
	Click Cloud Platform Project
	Enter the Project Number
	Click Set Project
	Hopefully, if you did everything correctly, you will get a confirmation message that it was successful
	
	Now you can close the Google Cloud Platform
	
	Show the Code.gs file
	Click the drop down list that states "Select function" and choose setGlobals
	Click run button, which is a button with a triangle
	You will be asked to authorize the script
	Authorize the permissions
