## Getting Started:
The Loop Package Generator is your solution for effortlessly automating the creation of commissioning loop packages. It seamlessly incorporates all the essential files using VBA, Python, and PyPDF2, making the process simple, fast, and highly effective.

**What to expect from the Loop Package Generator:** 

With one click you will have a LOOP PACKAGE containing:

Front Page + Record Sheets + Attachments + Punch List + P&ID 

All in one single PDF ready to use

## Set Up:
**Step 01:**

Download and install the latest version of Python from:
www.python.org/downloads

**Step 02:**

Confirm Python & PyPDF2 installation by clicking "Verify & Install PyPDF2"

**Note:** The PyPDF2 library will be automatically installed upon clicking the button.

![Verifygif](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/90efb118-8629-4db9-b1cb-8b38b1a4ef74)


**Step 03:**

Test the Loop Package Generator by clicking "Test Loop Package Generator"

Once that is done successfully you should see something like this:

![Untitled-6](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/59ae5f92-4348-4c43-957a-d548a13fc159)

A PDF will be generated and opened automatically with a template of a loop package folder

## How To Use:
**Content:** The Loop Package Generator will come with the following sheets

 - **LOOP PACKAGES:** This sheet contains information regarding each loop package
 - **Front Page:** This sheet serves as the cover of the loop package
 - **AI, AO, DI, DO:** These sheets provide examples of record sheets
 - **Punch List:** This sheet contains the punch list template

When you open the Excel file for the first time, the following folders will be created automatically:
- Attachments Folder, which will contain 10 subfolders and one for P&ID
- Loop Packages Folder, where the final Loop Packages will be stored when created

![image2](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/ebbc1890-866b-4399-9d77-05eae934bec2)

**Step 01:** Update Project Information 
Updating the project information is simple, just use the buttons provided

![image3](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/1a4eba8d-6231-47de-a856-25201ecb266e)


**Step 02:** Fill Out Loop Packages Sheet (An example is provided in the LOOP PACKAGES sheet)

 - **ID:** The ID of each loop package
 - **RESULT:** The output after running the Loop Generator can be "DONE" or "ERROR"
 - **LOOP PACKAGE INFORMATION:** Includes the loop SYSTEM No, SYSTEM DESCRIPTION, REPORT No, P&ID No, System Type.
 - **LOOP PACKAGE No:** The loop package number
 - **ASSOCIATED TAGS:** Up to 20 tags
 - **ATTACHMENTS**: Up to 10 attachments (PDF ONLY)

**Step 03:**
Copy your P&ID files and attachments to the newly created folders. For example, you can use Attachment 01 for your Instrument Data Sheet and Attachment 02 for Alarms Index, and so on

![image4](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/fbd54b03-413d-4363-adb5-a45b4a03ad45)

**Step 04:** Verify for Errors						
After filling out the information and copying the necessary files, it is recommended to verify your loops for errors and mistakes.

![image5](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/a7f0e5a7-8545-40b9-bb7a-4c11e5c65fa5)

Errors:						
 - **YELLOW COLOR:** Missing LOOP PACKAGE NUMBER or IO TYPE or TAG NAME
 - **RED COLOR:** IO TYPE Does Not Have a Record Sheet
 - **ORANGE COLOR**: File Not Found

**Step 05:** Create Loop Packages
After verification, click "Create Loop Packages" and check the Loop Packages Folder to find the results.

This is a real example of the Loop Package Generator in use: (+200 Folder was created under 3mn)

![image6](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/7cc91f6b-3ab6-4c4d-b184-553f947bcea5)

![image7](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/46a107b1-8331-4094-8610-238b7ec7f4a1)

**Step 06:** Creation of New Record Sheet
If you wish to add a new record sheet, follow these steps:
01. Click "CREATE NEW RECORD SHEET" in the LOOP PACKAGES SHEET
02. Write the record sheet name (ex, Digital Input, Analog Output) and type (ex, AI, AO, DI, DO).
03. Click "OK" to generate the new record sheet.

![image8](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/21f5babd-1872-4c49-a4ff-472ccfd99978)


## Important Information:
The Loop Package Generator works only with Windows for now.
