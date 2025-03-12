# Google-Apps-Script-project
A project I completed during my internship using Google Apps Script.

Maintaining communications is vital to having good relationships with donors and sponsors, so things like holiday thank-you letters that detail 
progress and goals for the new years, as well as expressing gratitude for their support is necessary. In this project I had to simply tailor each template thank you letter: adding the name and the mailing address of the recepient, as well as the date. I had to do this for well over 100 contacts, which would have been very tedious work and would take a lot of time, as well as increase the chances of human error. The information for each donor was in a pre-existing google sheet file. For privacy reasons, the data I have provided in the repository is not what I really used, but is fake data that is purely for viewing/demonstration. The template document referenced included is also for demonstration. Following is a step-by-step breakdown of the process to complete this project. 

1.) **Calling the function and the files**: After calling the function, which I named 'createletters', I define several constants pertaining to the necessary google products: the template document, reference sheet, and drive folder. They are defined by their unique file id's provided by Google. 

2.) **Accessing the files**: The sheet I want to use 'Sheet1', has to be specified and activated (selected for further processing) using the statement "spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Sheet1"));", for easier access, I created a constant to reference this sheet ("  const sheet = spreadsheet.getActiveSheet();"). After calling this statement "  const data = sheet.getDataRange().getValues();", the data is arranged as a 2D array that can be adequately worked with via javascript. The document and the folder are accessed with a variable declaration and method call. This was used to access the drive folder for example: "  const folder = DriveApp.getFolderById(folderid);"

3.) **For loop**: A loop variable is initialized with a starting variable of 1(i=1) in order to skip the header row. The for loop iterates over 'data' as long as the specified condition remains true (that i < data.length), I also added the increment step (i++) so each row adds 1 to 'i'. 

4.) **Extracting data from each row**: I define several important variables, the row number (I will use for naming each document; i+1 to skip the header row), the full name, first name, and address. The latter 3 I define via array indexing. For example, to retrieve the address, I define the 'address' variable as: "const address = data[i][1];", where the 'i' represents a row, and the 1 represents the 2nd column. The process repeats for the other two columns involved.

5.) **Creating a copy of the template document**: A few more variables have to be declared. A template title for each new document "${rowNumber} ${fullname} donation letter" is assigned to 'title'. A 'newdoc' variable is assigned a method chaining statement that retrieves the template document, makes a duplicate, a name, then it is sent to the folder; "DriveApp.getFileById(templatedoc).makeCopy(`${fullname} letter`, folder)". "newdocId" is assigned a new google file id that will later be referenced. The document has to be opened with an 'openById' call, then retrieve the text of the document in order to edit it with a 'getBody()' call. 

6.) **Replace placeholders**: Now I replace the placeholders in the copy of the template document, for the name placeholders, it follows a simple replace command: "body.replaceText("{{full name}}", fullname || "")", the OR operator replaces the placeholder with an empty string if there is no name. The address statement differs slightly in that I have to use a ternary operator because I am also using a replace command to break up the address down to two lines (("{{address}}", address ? address.replace(", ", "\n") : "")), so after the first occurence of a comma, there is a line break ('\n'). 

7.) **Saving and wrapping up**: The title template is applied to the finished document with "doc.setName(title)", and then changes are saved and the document is closed with "doc.saveAndClose();", then the loop is closed with the closing bracket. Finally a log statement indicates completion of the entire process, and the function is concluded with a closing bracket. 

Thank you for reading. 
   

