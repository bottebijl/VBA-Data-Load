Attribute VB_Name = "Buttons"
'Code behind the buttons on the START-sheet. Look up relevant procedures in module 'modEssbase or 'modLoadFile

Sub CreateINPUTSheet()
    Call prepare_sheets
End Sub

Sub RetrieveReportSht()
    Call RefreshReportSheet("REPORT_TEST", True, True)
    'First Variable is name of report sheet
    'First Boolean is to whether or not copy refreshed prior data to input sheet
    'Second Boolean is to ask whether or not to keep reference formula to report sheet _
    False means use values instead
End Sub

Sub ExportLoadfile()

    Call create_loadfile("INPUT_TEST")
    'First Variable is name of input sheet
End Sub

Sub CreateCheckFile()
    
    Call EssCreateCHECKReport
    
End Sub
