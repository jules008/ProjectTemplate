VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDatePicker 
   Caption         =   "Calendar"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   OleObjectBlob   =   "FrmDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub BtnClear_Click()
    On Error Resume Next
    Tag = ""
    Me.Hide
End Sub

Private Sub Calendar1_Click()
    On Error Resume Next
    Tag = Me.Calendar1.Day & "/" & Calendar1.Month & "/" & Calendar1.Year
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    Me.Calendar1.Day = Format(Now, "dd")
    Me.Calendar1.Month = Format(Now, "mm")
    Me.Calendar1.Year = Format(Now, "yyyy")
End Sub

