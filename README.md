This macro separates files based on unique identifier in a table. For example, if a table consist of different users Moe, Larry and Curly, after running the macro three files will be created with all the info only for Moe, Larry and Curly.

To use the Macro:
1.	Open the Excel sheet containing the master table with all the data. 
2.	Go to Developer tab and open Visual Basic. 
3.	In the new window “Microsoft Visual Basic for Applications”, Go to File -> Import File -> and select CopyNameRow.bas 

A new module under the “Modules” folder will be open. Double click it to open the visual basic code and enter the information under the user input 
- sMaster = 2:  Excel sheet that contains the master table (ex. Sheet #2)
- sReadme = 3: Readme tab. This tab will be placed as the first sheet in all the new sheets that will be created (ex. Sheet #3(
- tName = "Master":  Name of table that is being copied/pasted into different files (with "" around it). You can find this name by click on any field in the table and checking under “Table Design”
- colNum = 5: Column NUMBER which is being filtered i.e. where the unique ID such as manager names are.
- sFolder = "C:\Users\siqbal\VBA_Split_Files\Results\": Folder where all the spilt files are to be saved. This folder must be created manually prior to running the macro. 
