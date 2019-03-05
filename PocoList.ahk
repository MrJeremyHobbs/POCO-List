;Start;
FileDelete Poco_Inventory.docx
FileDelete _results_sorted.xlsx

IfNotExist, Template.docx
{
	msgbox Cannot find Template.docx
	exit
}

;Get spreadsheet path
IniRead, file_path, config.ini, spreadsheet, path

;Get input file
FileSelectFile, xlsFile,,%file_path%, Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile =
{
	exit
}

;Copy spreadsheet to program folder for editing
xlsFileCopy = _results_sorted.xlsx
FileCopy, %xlsFile%, %xlsFileCopy%, 1
xlsxFileCopyPath = %A_ScriptDir%\%xlsFileCopy%

;Status
Progress, zh0 fs12, Generating List...One Moment...,,Status

;Open XLS file
xl := ComObjCreate("Excel.Application")
xl.Visible := False
book := xl.Workbooks.Open(xlsxFileCopyPath)

;Get column references
IniRead, description_col, config.ini, columns, description_col
IniRead, title_col, config.ini, columns, title_col
IniRead, permanent_location_col, config.ini, columns, permanent_location_col

;Convert string references to integers
description_col := Round(description_col)
title_col := Round(title_col)
permanent_location_col := Round(permanent_location_col)

;Sort Columns (in reverse order)
;xlAscending = 1, xlYes = 1
;Permanent location is important because it sorts out the areas by check-out period.
xl.cells.sort(xl.columns(description_col), 1) ;Description
xl.cells.sort(xl.columns(title_col), 1) ;Title
xl.cells.sort(xl.columns(permanent_location_col), 1) ;Permanent Location


;Save and quit XLS file
book.Save()
book.Close
xl.Quit

;Open DOC file
template = %A_ScriptDir%\Template.docx
saveFile = %A_ScriptDir%\Poco_Inventory.docx
wrd := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform Mail Merge
doc := wrd.Documents.Open(template)
doc.MailMerge.MainDocumentType := 3 ;Mail merge type "directory"
doc.MailMerge.OpenDataSource(xlsxFileCopyPath,,,,,,,,,,,,,"SELECT * FROM [results$]")
doc.MailMerge.Execute

;Add header row
wrd.Selection.InsertRowsAbove(1)
wrd.Selection.Tables(1).Rows(1).Height := 30
wrd.Selection.Cells.VerticalAlignment := 1
wrd.Selection.ParagraphFormat.Alignment := 1
wrd.Selection.Shading.BackgroundPatternColor := -587137025
wrd.Selection.Font.Italic := False
wrd.Selection.Font.Bold := True
wrd.Selection.TypeText("Title")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Status")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Due Date/MSG")
wrd.Selection.Rows.HeadingFormat := 9999998 ;Set header for each page
    
;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFile)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Check for finished list
IfNotExist, Poco_Inventory.docx
{
	msgbox Cannot find Poco_Inventory.docx
	exit
}

;Clean-up
IniRead, delete_when_done, config.ini, misc, delete_when_done
if delete_when_done = active
{
    FileDelete %xlsFile%
}

;Finish
run winword.exe Poco_Inventory.docx
