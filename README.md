# Excel to XML
## Programs Used
Excel is a spreasheet application developed by Microsoft

VBA is an event-driven programming language for Microsoft Office applications

XML is a markup language without predefined tags 
## Purpose
Gather data from Excel and convert it to XML with only excel functions. 

The VBA code is used for gathering multiple data in a single cell and breaking it up into XML segments. 
## Who is this for
This technique of converting Excel data into XML is intended for non programmers. 
## Pros
1. Converts large Excel documents into XML by using simple Excel functions and calling to certain cells. 
2. Requires no external application 
## Cons 
1. Requires VBA coding solution when multiple data is in single cell 
## How to Convert With Excel Functions
1. Type XML header on top row of Excel
2. Use =CONCAT() function to collect data
3. Write out XML code calling to data in excel cell when needed
4. Type closing XML tags on last row
5. Copy entire XML cell into word doc and save as .txt file

Excel Function Example:

![alt text](https://github.com/LeoTech-Esploro/Excel_EsploroBatchEdit/blob/main/images/Excel_Functions_Example.jpg)
## How to Use VBA
1. Activate developer tab in Excel (Link for Instructions https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)
2. Click Visual Basic button in the Developer tab
3. Click on sheet1 and write VBA code that extracts multiple data from one cell by seperating data with key like ","
4. Hit the Run button at top

Example Code:

![alt_text](https://github.com/LeoTech-Esploro/Excel_EsploroBatchEdit/blob/main/images/Excel_VBA_Code.jpg)

Output:

![alt_text](https://github.com/LeoTech-Esploro/Excel_EsploroBatchEdit/blob/main/images/Excel_VBA_Output.jpg)
