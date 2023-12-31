Download the Loop Package Generator in the files section above or [Click Here](https://raw.githubusercontent.com/mokhtarbendaho/Loop-Package-Generator/main/Loop%20Package%20Generator%20V2.8.xlsm)

## Getting Started:
The Loop Package Generator is your solution for effortlessly automating the creation of commissioning loop packages. It seamlessly incorporates all the essential files using VBA, Python, and PyPDF2, making the process simple, fast, and highly effective.

**What to expect from the Loop Package Generator:** 

With one click you will have all your LOOP PACKAGES, each in a single PDF file containing a:

Front Page + Record Sheets + Attachments + Punch List + P&ID

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

**Step 04:** Update Project Information

Updating the project information is simple, just use the buttons provided (Only JPEG,JPG are accepted)

![image3](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/00979de3-a95d-4d8b-821b-dba4649cdb2a)

## Content:

The Loop Package Generator will come with the following sheets

 - **LOOP PACKAGES:** This sheet contains information regarding each loop package
 - **Front Page:** This sheet serves as the cover of the loop package
 - **AI, AO, DI, DO:** These sheets provide examples of record sheets
 - **Punch List:** This sheet contains the punch list template

When you open the Excel file for the first time, the following folders will be created automatically:
- Attachments Folder, which will contain 10 subfolders and one for P&ID
- Loop Packages Folder, where the final Loop Packages will be stored when created

![image2](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/ebbc1890-866b-4399-9d77-05eae934bec2)

## How To Use:

**Step 01:** Fill Out Loop Packages Sheet (An example is provided in the LOOP PACKAGES sheet)

 - **ID:** The ID of each loop package
 - **RESULT:** The output after running the Loop Generator can be "DONE" or "ERROR"
 - **LOOP PACKAGE INFORMATION:** Includes the loop SYSTEM No, SYSTEM DESCRIPTION, REPORT No, P&ID No, System Type.
 - **LOOP PACKAGE No:** The loop package number
 - **ASSOCIATED TAGS:** Up to 20 tags
 - **ATTACHMENTS**: Up to 10 attachments (PDF ONLY)

![image9](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/93e1c9c4-a7a5-4a9c-9644-91f930509365)


**Step 02:**
Copy your P&ID files and attachments to the newly created folders. For example, you can use Attachment 01 for your Instrument Data Sheet and Attachment 02 for Alarms Index, and so on

![image4](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/fbd54b03-413d-4363-adb5-a45b4a03ad45)

After copying the necessary files, fill out the Attachment section accordingly

**THE ATTACHMENT NAME IN THE EXCEL FILE SHOULD BE THE SAME AS THE FILE NAME IN THE ATTACHMENT FOLDERS, ONLY PDF FILES ARE ACCEPTED**

![image10](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/b0a9b7d7-073d-48ad-9b85-29dc8c2e97d1)


**Step 03:** Verify for Errors						
After filling out the information and copying the necessary files, it is recommended to verify your loops for errors and mistakes, by clicking on **"CHECK FOR ERRORS"**

![image5](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/a7f0e5a7-8545-40b9-bb7a-4c11e5c65fa5)

Errors:						
 - **YELLOW COLOR:** Missing LOOP PACKAGE NUMBER or IO TYPE or TAG NAME
 - **RED COLOR:** IO TYPE Does Not Have a Record Sheet
 - **ORANGE COLOR**: File Not Found

**Step 04:** Create Loop Packages
After verification, click **"CREATE LOOP PACKAGES"** and check the Loop Packages Folder to find the results.

This is what will you see as a result:

![image6](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/f6531380-b885-4df6-9137-ee4ed9ae62e5)

![image7](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/57bc013d-9d71-463f-a264-c06d5a4524bc)

Also, by checking the following option on the main page:

![image](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/8552dd39-2665-49a7-bd1b-c38c0eee6b22)

 You have the option to add a blank A3 page at the end if there is no P&ID, this will be useful when printing in bulk, it will make it easier to separate the folders

**Step 05:** Creation of New Record Sheet

If you wish to add a new record sheet, follow these steps:
01. Click **"CREATE NEW RECORD SHEET"** in the LOOP PACKAGES SHEET.
02. Write the record sheet name (ex, Digital Input, Analog Output, FieldBus).
03. Write the IO type (ex, AI, AO, DI, DO, FB).
04. Click "OK" to generate the new record sheet.

![image8](https://github.com/mokhtarbendaho/Loop-Package-Generator/assets/143171867/21f5babd-1872-4c49-a4ff-472ccfd99978)

## Conclusion:

I trust that the Loop Package Generator will prove as beneficial to you as it has been for my team and me during our commissioning activities.

If you have any questions or need any help feel free to contact me.

Email: bendahomokhtar@gmail.com

LinkedIn: www.linkedin.com/in/mokhtar-bendaho/

**Important Information:** The Loop Package Generator works only with Windows for now.
**Disclaimer:** The record sheets in this workbook are for illustration only and should not be used without prior client approval and legal compliance. By using this Excel file, you agree to our [terms and conditions](https://github.com/mokhtarbendaho/Loop-Package-Generator/blob/main/Terms%20and%20Conditions).
