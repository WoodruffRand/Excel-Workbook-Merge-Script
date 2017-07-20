# Excel-Workbook-Merge-Script
Merges sheet in open active work book into workbook macro resides in
Loads in net data to excel sheet from outside work book
desiged speciafly to add new data to ongoing excel data repositroy's (for when you cant convice your boss to use a database)
Place maco in "data base" and then run on sheet you intend to merge.  

assumptions:
1.  workbook to import must be active sheet
2.  workbook to recive file must contain the marco
3.  data has labels starting on row 2 of "database" file and line 1 of input file
4.  varriable names are identical, any non matched data field will not be imported.
5.  runs in n^2(columns) time, so if you have 1k+ columns, some opptimatizion would help.

Script will only import matching data and copy it to appropriate location
so this is uslefull when shape of inital data has changed from intal source 
