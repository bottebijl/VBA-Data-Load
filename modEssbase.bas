Attribute VB_Name = "modEssbase"
Public MyNote, Answer As String
Public varAccounts, varLocations, varCountLoc, varRow As Integer
Public first_datarow, last_datarow As Long
Public last_row As Integer
Public col_org_index, col_acc_index As Integer
Public target_sht As Worksheet
Public data_rng As String
Public start_sht, input_sht, report_sht, check_sht, delta_sht As String

Sub SetVariables()

'Set the names of sheets for use in functions below
start_sht = ThisWorkbook.Sheets("START").Name
input_sht = "" & ThisWorkbook.Sheets("INPUT_TEST").Name
report_sht = "" & ThisWorkbook.Sheets("REPORT_TEST").Name
check_sht = ThisWorkbook.Sheets("CHECK_TEST").Name
delta_sht = ThisWorkbook.Sheets("DELTA_TEST").Name

'nr of Accounts/Location types for particular FDMEE Location
varAccounts = Range("nrAccounts")
varLocations = Range("nrLocations")

col_org_index = 4
col_acc_index = 3

'Set first and last datarow as range to work with (60000 chosen as highest number possible)
first_datarow = 17
last_datarow = 60000
End Sub

Sub prepare_sheets()

MyNote = "Are you sure? - this will rebuild the INPUT sheet and delete all data or mappings in the INPUT sheet"
Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Create New INPUT-sheet")

If Not Answer = vbYes Then End 'Exit when users do not confirm

application.ScreenUpdating = False

Call SetVariables

'Set up all relevant worksheets
Call prepare_sht(input_sht)
Call prepare_sht(report_sht)
Call prepare_sht(check_sht)
Call prepare_sht(delta_sht)

'Hide Worksheets not relevant for users
Call ZunhideAll
Call Zhideall

application.ScreenUpdating = True

MsgBox "Successful"

End Sub

Sub prepare_sht(ByVal sht As String)
          

Set target_sht = ThisWorkbook.Sheets(sht)
     
Dim arr()
       
ReDim Preserve arr(0 To (varLocations * varAccounts), 2)
arr_row = 0
        
    varCountLoc = 1
    For j = 1 To varLocations
        For i = 1 To varAccounts
                
            arr(arr_row, 1) = j
            arr(arr_row, 2) = i
                
            arr_row = arr_row + 1
        Next
            varCountLoc = varCountLoc + varAccounts
        
    Next
        
        With target_sht
            
            'Delete all data rows
            .Rows(first_datarow & ":" & last_datarow).Delete
            
            'Set the Location and Account indexes in respective columns
            .Range(target_sht.Cells(first_datarow, col_acc_index), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), col_acc_index)).Value = application.Index(arr, , 3)
            .Range(target_sht.Cells(first_datarow, col_org_index), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), col_org_index)).Value = application.Index(arr, , 2)
            
            'Find dimension members based on given Account & Location indexes
            .Range(target_sht.Cells(first_datarow, 24), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 24)).FormulaR1C1 = "=if(index(arLocation," & "RC[-20]" & ",indexLocations+2)=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",index(arLocation," & "RC[-20]" & ",indexLocations+2))"
            .Range(target_sht.Cells(first_datarow, 25), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 25)).FormulaR1C1 = "=if(index(arAccounts," & "RC[-22]" & ",indexAccounts)=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",index(arAccounts," & "RC[-22]" & ",indexAccounts))"
            .Range(target_sht.Cells(first_datarow, 26), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 26)).FormulaR1C1 = "=if(index(arLocation," & "RC[-22]" & ",indexLocations)=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",index(arLocation," & "RC[-22]" & ",indexLocations))"
            
            If target_sht.Name = delta_sht Then 'Include delta checks only for delta sheet
                'Include substraction formula in data fields to compare input sheet with check sheet
                .Range(target_sht.Cells(first_datarow, 27), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 38)).FormulaR1C1 = "=" & check_sht & "!RC-" & input_sht & "!RC"
                .Range(target_sht.Cells(first_datarow - 4, 27), target_sht.Cells(first_datarow - 4, 39)).FormulaR1C1 = "=SUM(R[+4]C:R[" & UBound(arr, 1) + 3 & "]C)"
            End If
          
            'Add totals to row
            .Range(target_sht.Cells(first_datarow, 39), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 39)).FormulaR1C1 = "=SUM(RC[-12]:RC[-1])"
            
            'Show user friendly names for location+loc type and Account in first to columns
            .Range(target_sht.Cells(first_datarow, 1), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 1)).FormulaR1C1 = "=if(index(arLocation," & "RC[+3]" & ",indexLocations+1)=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",index(arLocation," & "RC[+3]" & ",indexLocations+1) & "" - "" & index(arLocation," & "RC[+3]" & ",indexLocations+2))"
            .Range(target_sht.Cells(first_datarow, 2), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 2)).FormulaR1C1 = "=if(index(arAccounts," & "RC[+1]" & ",indexAccounts+1)=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",index(arAccounts," & "RC[+1]" & ",indexAccounts+1))"
            
            'Copies the (Conditional) Format from row 12 to the entire range for data
            .Range(target_sht.Cells(12, 27), target_sht.Cells(12, 39)).Copy
            .Range(target_sht.Cells(first_datarow, 27), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 39)).PasteSpecial xlPasteFormats
                        
            .Outline.ShowLevels rowlevels:=1, columnlevels:=1
           
        End With
    
If target_sht.Name = input_sht Then
    ThisWorkbook.Sheets("ADMIN").Range("setDataRange") = target_sht.Range(target_sht.Cells(first_datarow, 27), target_sht.Cells(first_datarow - 1 + UBound(arr, 1), 38)).Address
End If

End Sub

Sub RefreshReportSheet(vbReportSheet As String, vbCopyToInput As Boolean, vbKeepFormula As Boolean)
      
Call SetVariables
    
    Dim Answer As String
    Dim MyNote As String
    Dim varRows As Long
    Dim lastRow As Long
    Dim mysheet As String
    Dim start_column As Long
    Dim last_row As Long
    Set target_sht = ThisWorkbook.Sheets(input_sht)
    
    'Message for messagebox
    MyNote = "Retrieve the current Forecast from Essbase and fill INPUT sheet?"
            
    'Display MessageBox
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "OPE G&A reporting template")
    
    
    
    If Answer = vbYes Then ' only execute when users wants to proceed
    application.ScreenUpdating = False
        Calculate
        ThisWorkbook.Sheets(vbReportSheet).Select
        
        'Check if Smart View is installed
        sts = 10
        'Get Smartview Version
        On Error Resume Next
        sts = HypGetVersion(BUILD_VERSION, Version, 0)
        'MsgBox "Smartview version: " & Version
        If sts = 0 Then
            'Run Smartview report is Smartview is installed
            Call vbaRetrieve("REPORT_TEST")
        
        End If
      
        'Group rows and columns again
        ActiveSheet.Outline.ShowLevels rowlevels:=0, columnlevels:=1
        
        If vbCopyToInput Then 'if copy to inputsheet is set as true in calling routine
            
            start_column = Range("setStartColumn").Value
            last_row = ThisWorkbook.Sheets(report_sht).Range("A60000").End(xlUp).row
            
            With target_sht 'perform action on sheet set as input sheet
                
                If vbKeepFormula Then
                    .Range(target_sht.Cells(17, start_column), target_sht.Cells(last_row, 38)).FormulaR1C1 = "=" & report_sht & "!RC"
                Else
                    .Range(target_sht.Cells(17, start_column), target_sht.Cells(last_row, 38)).FormulaR1C1 = "=" & report_sht & "!RC"
                    .Range(target_sht.Cells(17, start_column), target_sht.Cells(last_row, 38)).Copy
                    .Range(target_sht.Cells(17, start_column), target_sht.Cells(last_row, 38)).PasteSpecial (xlPasteValues)
                End If
             
            End With
                 
        End If
        
        ThisWorkbook.Sheets("START").Select
        
        application.ScreenUpdating = True
        MsgBox ("Data is retrieved from Essbase")
    End If
End Sub
Sub EssCreateCHECKReport()
    
Call SetVariables
    
Dim CheckAnswer As String
Dim CheckMyNote As String
    
Set target_sht = ThisWorkbook.Sheets(check_sht)
    
CheckMyNote = "Is the GA data uploaded with Workspace to Essbase?"

'Display MessageBox
CheckAnswer = MsgBox(CheckMyNote, vbQuestion + vbYesNo, "OPE G&A reporting template")
application.ScreenUpdating = False
    
If CheckAnswer = vbYes Then
        
    target_sht.Select
        
    'Check if Smart View is installed
    sts = 10
    'Get Smartview Version
    On Error Resume Next
    sts = HypGetVersion(BUILD_VERSION, Version, 0)
    
    If sts = 0 Then
        'Run Smartview report is Smartview is installed
        Call vbaRetrieve(target_sht.Name)
    End If
      
    ActiveSheet.Outline.ShowLevels rowlevels:=0, columnlevels:=1
    
    ThisWorkbook.Sheets(start_sht).Select
    application.ScreenUpdating = True
        
    MsgBox ("Data is retrieved from Essbase")
    
End If

End Sub

Sub ZunhideAll()
    Sheets("ADMIN").Visible = True
    Sheets("LOAD").Visible = True
    Sheets("ACCOUNTS").Visible = True
    Sheets("LOCATIONS").Visible = True
    Sheets("START").Select
    Range("D5").Select
End Sub

Sub Zhideall()
    Sheets("ADMIN").Visible = False
    Sheets("LOAD").Visible = False
    'Sheets("ACCOUNTS").Visible = False
    'Sheets("LOCATIONS").Visible = False
    Sheets("START").Select
    Range("D5").Select
End Sub

