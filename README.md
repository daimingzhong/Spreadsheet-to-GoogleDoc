# Google-Spreadsheet-to-Doc

## Instructions: 
### Request and Running Example
New google doc files are generated under the same folder of your spreadsheet. So it is better to place you spreadsheet in one folder.
For running the script, you will need to setup three files.
1. A spreadsheet, which contains information for generating docs. Note that one line record will generate one doc.
2. A template google doc file named as GoogleDocTemplate.gdoc
3. A template spreadsheet named as GoogleSheetTemplate.gsheet, which is used for saving key-value pairs and new file name.

For instance, you have line 10th recording studentID information and line 15th recording age.(counting columns from 0), and you are expecting to have that information in new google doc.  
The operations are as follows: 
1. Put << student id>> and << age>> in the proper place of GoogleDocTemplate. 
2. In GoogleSheetTemplate, write the first 4 cells as << student id>> 10 << age>> 15, one line corresponds to one mapping relationship.  
3. Write the new docs' name in the first cell of second page of GoogleSheetTemplate(as sheet2). You may write as "studentID is << student id>> and age is << age>>", then you may have the file name as
"studentID is 2015015 and age is 12".

### Time Stamp
1. You can use Recent_Date to specify the format with MM/dd/yyyy.

## References
Many thanks to Mikko Ohtamaa, http://opensourcehacker.com  
https://opensourcehacker.com/2013/01/21/script-for-generating-google-documents-from-google-spreadsheet-data-source/  

Youtube tutorials:  
https://www.youtube.com/watch?v=QKoCltmJPvs&list=PL8TcZ-5gvt5jmLRaA4Xfd2GSDKHncbAYm&index=7  

https://www.youtube.com/watch?v=5_8T1KKboEU


