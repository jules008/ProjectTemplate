VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsCalButton
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 29 Apr 20
'===============================================================

Option Explicit

Public WithEvents CmdBtnGroup As MSForms.CommandButton
Attribute CmdBtnGroup.VB_VarHelpID = -1

Sub CmdBtnGroup_Click()

    If Month(CDate(CmdBtnGroup.Tag)) <> FrmCalPicker.CmoMonth.ListIndex + 1 Then
        Select Case _
               MsgBox("The selected date is not in the currently selected month." _
                      & vbNewLine & "Continue?", _
                      vbYesNo Or vbExclamation Or vbDefaultButton1, "Date check")
            Case vbYes
                If G_FORM Then
                    GoTo on_Form
                Else: GoTo AddDate
                End If
            Case vbNo
                Exit Sub
        End Select
    Else:
    If G_FORM Then
        GoTo on_Form
    Else: GoTo AddDate
    End If
AddDate:

    FrmCalPicker.ReturnDate = CDate(CmdBtnGroup.Tag)

    GoTo chg_month
on_Form:
    G_DATE = CmdBtnGroup.Tag
chg_month:
    With FrmCalPicker.CmoMonth
        .ListIndex = Month(CmdBtnGroup.Tag) - 1
    End With
     End If
    Unload FrmCalPicker
End Sub
