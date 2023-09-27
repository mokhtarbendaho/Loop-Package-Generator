## Getting Started:
The Loop Package Generator is your solution for effortlessly automating the creation of commissioning loop packages. It seamlessly incorporates all the essential files using VBA, Python, and PyPDF2, making the process simple, fast, and highly effective.

## Important Information:
The Loop Package Generator works only with Windows for now.

## Set Up:
**Step 01:**
Download and install the latest version of Python from:

www.python.org/downloads

<p>&nbsp;</p>

**Step 02:**

Confirm Python & PyPDF2 installation by clicking "Verify & Install PyPDF2"

**Note:** The PyPDF2 library will be automatically installed upon clicking the button.

<p>&nbsp;</p>

**Step 03:**
Test the Loop Package Generator by clicking "Test Loop Package Generator"

Once that is done you should see something like this:



## How To Use:
**Content:** The Loop Package Generator will come with the following sheets

 - **LOOP PACKAGES:** This sheet contains information regarding each loop package
 - **Front Page:** This sheet serves as the cover of the loop package
 - **AI, AO, DI, DO:** These sheets provide examples of record sheets
 - **Punch List:** This sheet contains the punch list template

When you open the Excel file for the first time, the following folders will be created automatically:
- Attachments Folder, which will contain 10 subfolders and one for P&ID
- Loop Packages Folder, where the final Loop Packages will be stored when created

**Step 01:** Update Project Information 
Updating the project information is simple, just use the buttons provided




**Step 02:** Fill Out Loop Packages Sheet (An example is provided in the LOOP PACKAGES sheet)

 - **ID:** The ID of each loop package
 - **RESULT:** The output after running the Loop Generator can be "DONE" or "ERROR"
 - **LOOP PACKAGE INFORMATION:** Includes the loop SYSTEM No, SYSTEM DESCRIPTION, REPORT No, P&ID No, System Type.
 - **LOOP PACKAGE No:** The loop package number
 - **ASSOCIATED TAGS:** Up to 20 tags
 - **ATTACHMENTS**: Up to 10 attachments (PDF ONLY)

**Step 03:**
Copy your P&ID files and attachments to the newly created folders. For example, you can use Attachment 01 for your Instrument Data Sheet and Attachment 02 for Alarms Index, and so on

**Step 04:** Verify for Errors						
After filling out the information and copying the necessary files, it is recommended to verify your loops for errors and mistakes.

**Step 05:** Create Loop Packages
After verification, click "Create Loop Packages" and check the Loop Packages Folder to find the results.

**Step 06:** Creation of New Record Sheet
If you wish to add a new record sheet, follow these steps:
01. Click "CREATE NEW RECORD SHEET"
02. Write the record sheet name (ex, Digital Input, Analog Output) and type (ex, AI, AO, DI, DO).
03. Click "OK" to generate the new record sheet.
