# Google-Apps-Script-project
A project I completed during my internship using Google Apps Script.

Maintaining communications is vital to having good relationships with donors and sponsors, so things like holiday thank-you letters that detail 
progress and goals for the new years, as well as expressing gratitude for their support is necessary. In this project I had to simply tailor each thank you letter: adding the name and the mailing address of the recepient, as well as the date. I had to do this for well over 100 letters, which would have been very tedious work and would take a lot of time, as well as increase the chances of human error. The information for each donor was in a pre-existing google sheet file. For privacy reasons, the data I have provided in the repository is not what I really used, but is fake data that is purely for viewing/demonstration. The template document referenced included is also for demonstration. Following is a step-by-step breakdown of the process to complete this project. 

1.) **Calling the function and the files**: After calling the function, which I named 'createletters', I define several constants pertaining to the necessary google products: the template document, reference sheet, and drive folder. They are defined by their unique file id's provided by Google. 

2.) **Accessing the files**: The sheet I want to use 'Sheet1', has to be specified and activated (selected for further processing) using the statement "spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Sheet1"));", for easier access, I created a constant to reference this sheet ("  const sheet = spreadsheet.getActiveSheet();"). After calling this statement "  const data = sheet.getDataRange().getValues();", the data is arranged as a 2D array that can be adequately worked with via javascript. The document and the folder are accessed with a variable declaration and method call. This was used to access the drive folder for example: "  const folder = DriveApp.getFolderById(folderid);"

3.) **For loop**: A loop variable is intitialized with a starting variable of 1(i=1) in order to skip the header row. The for loop iterates over 'data' as long as the specified condition remains true (that i < data.length), I also added the increment step (i++) so each row adds 1 to 'i'. 

4.) **extracting data from each row**: 
   

