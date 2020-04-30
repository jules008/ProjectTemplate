VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTimePicker 
   Caption         =   "Enter Time"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   OleObjectBlob   =   "FrmTimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmTimePicker
' displays form to select time
'---------------------------------------------------------------
' Created by Julian Turner
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 29 Apr 20
'===============================================================

Private Const StrMODULE As String = "ModCloseDown"

Option Explicit

' ===============================================================
' BtnClear_Click
' ---------------------------------------------------------------
Private Sub BtnClear_Click()
    TxtTime = "hh:mm"
    
        With TxtTime
            .SetFocus
            .SelStart = 0
            .SelLength = Len(TxtTime)
        End With
End Sub

' ===============================================================
' BtnEnterTime_Click
' ---------------------------------------------------------------
Private Sub BtnEnterTime_Click()
    If Not IsTime(TxtTime) Then
        With TxtTime
            .SetFocus
            .SelStart = 0
            .SelLength = Len(TxtTime)
        End With
        MsgBox "invalid time"
    Else
        MsgBox "time is " & TxtTime
    End If

End Sub

' ===============================================================
' TxtTime_AfterUpdate
' ---------------------------------------------------------------
Private Sub TxtTime_AfterUpdate()
    TxtTime = Format(TxtTime, "hh:mm")
End Sub

' ===============================================================
' TxtTime_Change
' ---------------------------------------------------------------
Private Sub TxtTime_Change()
    If Len(TxtTime) = 2 Then TxtTime = TxtTime & ":"
End Sub

' ===============================================================
' UserForm_Activate
' ---------------------------------------------------------------
Private Sub UserForm_Activate()
    TxtTime = Format("hh:mm", "hh:mm")
    With TxtTime
        .SetFocus
        .SelStart = 0
        .SelLength = Len(TxtTime)
    End With
End Sub
