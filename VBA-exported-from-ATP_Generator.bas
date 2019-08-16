Attribute VB_Name = "Module1"
Sub CloseBook()
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub

Sub resetthings()
    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> "START" Then
            xWs.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Sub RUN()
Attribute RUN.VB_ProcData.VB_Invoke_Func = " \n14"
'
'TO DO:

'Select file: excel file.
'copy stuff from that excel file to this one.
'select directory to save the _CSV_ file to.
'make the entire spreadsheet a table
'create 3 new columns.
'enter in the 3 equations to isolate ch, sect, item
'Remove the columns that don't matter. -- keep ch/sect/item and the CGI
'save the doc.
'close excel.


'create IPS Export tab
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = "sheet1"
    Sheets("sheet1").Select
    
        
        
        
  'PICK THE FILE
  msgbox "Please select the adhoc item-info-cgi File."
  
   'Declare a variable as a FileDialog object.
 Dim fd As FileDialog

 'Create a FileDialog object as a File Picker dialog box.
 Set fd = Application.FileDialog(msoFileDialogFilePicker)

 'Declare a variable to contain the path
 'of each selected item. Even though the path is aString,
 'the variable must be a Variant because For Each...Next
 'routines only work with Variants and Objects.
 Dim vrtSelectedItem As Variant

 'Use a With...End With block to reference the FileDialog object.
 With fd

 'Allow the user to select multiple files.
 .AllowMultiSelect = False

 'Use the Show method to display the File Picker dialog box and return the user's action.
 'If the user presses the button...
 If .Show = -1 Then
 'Step through each string in the FileDialogSelectedItems collection.
 For Each vrtSelectedItem In .SelectedItems



 'new variables as the two workbooks (IPS report and this one)
 Dim WBCalculator As Workbook, wb1 As Workbook

'we are officially setting This Workbook as WBCalculator
Set WBCalculator = ThisWorkbook

'And we're opening the file we selected as the IPS export (the vrtSelectedItem from the picker) as wb1
Set wb1 = Workbooks.Open(vrtSelectedItem)

'Copy Data from Wb1.Sheet1 to WBCalculator IPS Export sheet
wb1.Sheets(1).Range("A1").CurrentRegion.Copy WBCalculator.Worksheets("sheet1").Range("A1")
    Range("A2").Select
    
    
'close the ips export file, no alerts.  then turn alerts back on

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    wb1.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    
  ' rest of the loop thing
     Next
 'If the user presses Cancel...
 Else
 msgbox "you need to select the item info cgi adhoc file.  Please try again."
 'remove any tabs that have been created and reset to start
 
resetthings
 Exit Sub
 End If
 End With

 'Set the object variable to Nothing.
 Set fd = Nothing
    
    





'ActiveSheet.Name = "sheet1"


Sheets("sheet1").Select

'if the data has a filter - remove the filter.
If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
'if the data was a table - remove the table.
If ActiveSheet.ListObjects.Count > 0 Then ActiveSheet.ListObjects(1).Unlist

    Dim tbl As ListObject
'HeadersBoolean = msgbox("Does this spreadsheet have headers?", vbYesNoCancel)










  '  Else
    'hit cancel or x
 '   msgbox "Program exited."
  '  Exit Sub
'End If


'make the table


ActiveSheet.Range("a1:" & _
   ActiveSheet.Range("a1").End(xlToRight).End(xlDown).Address).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "sheet1"
    tbl.TableStyle = ""

'Sort the table



tbl.ListColumns.Add(3).Name = "item"
tbl.ListColumns.Add(3).Name = "sect"
tbl.ListColumns.Add(3).Name = "ch"

'add chapter
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=LEFT(SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"),FIND(""_"",SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"))-1)"
    Range("C3").Select

'add section
    Range("D2").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(LEFT(SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"),FIND(""."",SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"))-1),FIND(""_"",SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"))+1,LEN(SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """")))"
    Range("D3").Select


'add item
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=MID(LEFT(SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"),FIND(""m"",SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"))-1),FIND(""."",SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """"))+1,LEN(SUBSTITUTE(SUBSTITUTE([@[item-name]],[@[product-code]], """"), ""_ch"", """")))"
    Range("E3").Select


'copy all table and paste as values to remove formulas

    Range("sheet1").Select
    Range("C7").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False



Dim Output As String

'where do you want it output
msgbox "Please select where you would like the output file to be saved."


'use the get folder function
outputfolderlocation = GetFolder

'msgbox outputfolderlocation
If IsEmpty(outputfolderlocation) Then
msgbox "You need to select the folder location you would like the output folder to be saved.  Please try again."
resetthings
Exit Sub
Else
'Output = outputfolderlocation & "\adf.xml"

  '   Dim wb As Workbook
  '  Set wb = Workbooks.Add
  '  ThisWorkbook.Sheets("Sheet1").Copy Before:=wb.Sheets(1)
  '  wb.SaveAs outputfolderlocation & "\adf.xlsx"
  
  
  
  
  
    Sheets("sheet1").Select
    Sheets("sheet1").Move
   ' ChDir "C:\Users\sfrantz\OneDrive - Cengage Learning\Desktop"
    ActiveWorkbook.SaveAs Filename:=outputfolderlocation & "\adf.xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
   ' Range("H12").Select
   ' Windows("ATP_Generator.xlsm").Activate
   
   Workbooks("adf.xlsx").Close SaveChanges:=False
   ThisWorkbook.Close False
    
End If












End Sub


Sub RESET1()
'
' RESET1 Macro
'

    'delete the export column
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    

'if the data has a filter - remove the filter.
If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
'if the data was a table - remove the table.
If ActiveSheet.ListObjects.Count > 0 Then ActiveSheet.ListObjects(1).Unlist

End Sub





Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function
