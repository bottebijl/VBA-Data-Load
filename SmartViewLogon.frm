VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SmartViewLogon 
   Caption         =   "Smartview Logon"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4110
   OleObjectBlob   =   "SmartViewLogon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SmartViewLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Cancel_Click()
    Unload Me
    SmartViewLogon.varCancel.Value = True
End Sub
Private Sub OK_Click()
        
    Me.Hide
    SmartViewLogon.varCancel.Value = False
    
End Sub

