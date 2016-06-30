Attribute VB_Name = "modLoadFile"
Sub create_loadfile(vbInputSheet As String)

'dim txt file variables
Dim outputfile As Object
Dim fso As New Scripting.FileSystemObject

sFileName = application.GetSaveAsFilename(Range("setFilename").Value, "Text Files,*.txt") 'get_file_name & ".txt"
Set outputtextfile = fso.CreateTextFile(sFileName, True)

'Dim variables used in routine
Dim row_arr As Variant
Dim col_arr As Variant
Dim input_rng As Range
Dim input_arr As Variant
Dim sStr As String
Dim data_row As Integer
Dim data_col As Integer
Dim report_zeros As Boolean
Dim row_count As Integer

'set arrays for identifying where dimension members are stored in input sheet (for rows and columns)
row_arr = dim_array(True)
col_arr = dim_array(False)

'write entire input sheet (range defined in ADMIN sheet) to array
Set input_rng = ThisWorkbook.Sheets(vbInputSheet).Range(ThisWorkbook.Sheets("ADMIN").Range("setInputRange").Value)
input_arr = input_rng

'create empty variable for rowcount
row_count = 0

'boolean to determine whether to report zeros (set in ADMIN sheet)
report_zeros = ThisWorkbook.Sheets("Admin").Range("fdmReportZeros").Value

'loop through dimension members in rows on input sheet
For data_row = LBound(input_arr, 1) To UBound(input_arr, 1)
    
    'check if row is a data row by validating if the cells for row dimension members are filled
    filled_row = True
    For i = LBound(col_arr) To UBound(col_arr)
        If Not input_arr(data_row, col_arr(i)) <> "" Then
            filled_row = False
        End If
    Next i
        
    If filled_row = False Then 'do nothing if dimension members are empty
        
    Else 'row is data row, will be used for writing to text file
        
        For data_col = start_period_position(input_arr) To end_period_position(input_arr) '(UBound(input_arr, 2) - 11) To (UBound(input_arr, 2))
            sStr = ""
            If (input_arr(data_row, data_col) = "" Or input_arr(data_row, data_col) = 0) And Not report_zeros Then 'if not filled, fill with HashMissing or ignore (depending on 'fdmReportZeros' True or False
                GoTo NextIteration
            Else
            
            'Add  dimension members defined in columns to string
            For i = LBound(col_arr) To UBound(col_arr)
                sStr = sStr & input_arr(data_row, col_arr(i)) & ","
            Next i
            
            'Add  dimension members defined in rows to string
            For i = LBound(row_arr) To UBound(row_arr)
                sStr = sStr & input_arr(row_arr(i), data_col) & ","
            Next i
            
            If input_arr(data_row, data_col) = 0 Or input_arr(data_row, data_col) = "" Then '= "" Then 'if data = 0 then fill in #HashMissing
                sStr = sStr & "#HashMissing"
            Else
                sStr = sStr & data_calculation((input_arr(data_row, data_col)), input_arr, data_row, data_col) 'if filled, data value is added to end of line
            End If
            
            row_count = row_count + 1
            outputtextfile.WriteLine sStr 'write string to txt file
        
            End If

NextIteration:
        
        Next data_col 'move to next reporting period
    
    End If
    
Next data_row 'move to next account (row in input sheet)

'close file and terminate variables
outputtextfile.Close
Set outputtextfile = Nothing
Set fso = Nothing

MsgBox row_count & " lines have been loaded to load file.", vbInformation, "Saving loadfile successfull"

End Sub

Function get_file_name() As String

With application.FileDialog(msoFileDialogSaveAs)
    .Title = "Select where file needs to be stored"
    .AllowMultiSelect = False
    .InitialFileName = ThisWorkbook.Sheets("ADMIN").Range("setFilename")
    .Show
If .SelectedItems.Count = 0 Then
    End
ElseIf .SelectedItems.Count > 0 Then
    get_file_name = .SelectedItems(1)
End If

End With
End Function

Function dim_array(row_bool As Boolean) As Variant

Dim arr As Variant
Dim row_col As String
Dim T_row As Range
Dim index_nr As Integer

If row_bool Then row_col = "row" Else row_col = "col" 'if row_bool is true, then all row values for in dim_settings table are stored to array, otherwise all col values are stored

For Each T_row In [dim_settings].Rows
            
    If (Intersect(T_row, [dim_settings[Type]]).Value = "Dimension") And (Intersect(T_row, [dim_settings[row/col]]).Value = row_col) Then
            
        index_nr = Intersect(T_row, [dim_settings[Value]]).Value
            
        If IsEmpty(arr) Then
            
            arr = Array(index_nr) 'create array if not yet created
        
        Else
        
            ReDim Preserve arr(UBound(arr) + 1) 'if array exists in memory, redim it and create one more slot for next value to store in it
            arr(UBound(arr)) = index_nr 'append value to array
        
        End If
    
    End If

Next T_row

dim_array = arr

End Function


Function start_period_position(arr As Variant) As Integer

Dim StartPeriod As Integer

StartPeriod = 1 * (Right(ThisWorkbook.Sheets("ADMIN").Range("povPeriod").Value, 2))

start_period_position = (UBound(arr, 2) - (12 - StartPeriod))

End Function


Function end_period_position(arr As Variant) As Integer

Dim StartPeriod As Integer
Dim MultiLoad As Boolean

StartPeriod = Int(Right(ThisWorkbook.Sheets("ADMIN").Range("povPeriod").Value, 2))
MultiLoad = ThisWorkbook.Sheets("Admin").Range("fdmMultiLoad").Value

If Not MultiLoad Then
    end_period_position = (UBound(arr, 2) - (11 - StartPeriod))
Else
    end_period_position = (UBound(arr, 2))
End If

End Function

Function data_calculation(tgt As Double, arr As Variant, cur_row As Integer, cur_col As Integer) As String
'This function can be used to manipulate outcome of data value. At time of writing Scaling was the only
'manipulation required. This function can be expanded if other calculations need to be performed

Dim scaling As Integer

For Each T_row In [dim_settings].Rows
            
    If (Intersect(T_row, [dim_settings[Type]]).Value = "Function") And (Intersect(T_row, [dim_settings[Variable]]).Value = "Scaling") Then
        
        Dim specified_value As Integer
        specified_value = Intersect(T_row, [dim_settings[Value]]).Value
            
        If Intersect(T_row, [dim_settings[row/col]]).Value = "row" Then
            scaling = arr(specified_value, cur_col)
        ElseIf Intersect(T_row, [dim_settings[row/col]]).Value = "col" Then
            scaling = arr(cur_row, specified_value)
        Else 'do nothing
        End If
    End If

Next T_row
    

data_calculation = Format((tgt / scaling), "General Number")

'applicaton.UseSystemSeparators = True

End Function



