VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCalPicker 
   Caption         =   "Calendar"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3810
   OleObjectBlob   =   "FrmCalPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCalPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Module FrmCalPicker
'===============================================================
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 29 Apr 20
'===============================================================
Private Const StrMODULE As String = "FrmCalPicker"

Dim Buttons() As New ClsCalButton
Public ReturnDate As Date

Public Function ShowForm() As Date
    Dim CmdBtns As Integer
    Dim Cntrl As Control

    CmdBtns = 0
    For Each Cntrl In FrmCalPicker.Controls
        If TypeName(Cntrl) = "CommandButton" Then
            If Cntrl.Name <> "CB_Close" Then
                CmdBtns = CmdBtns + 1
                ReDim Preserve Buttons(1 To CmdBtns)
                Set Buttons(CmdBtns).CmdBtnGroup = Cntrl
            End If
        End If
    Next Cntrl

    FrmCalPicker.Show
    ShowForm = ReturnDate
End Function

Private Sub BtnClose_Click()
    Unload Me
End Sub

Sub AddDate()
    ReturnDate = Parent
End Sub

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim lYearsAdd As Long
    Dim lYearStart As Long

    lYearStart = Year(Date) - 10
    lYearsAdd = Year(Date) + 10
    With Me
        For i = 1 To 12
            .CmoMonth.AddItem Format(DateSerial(Year(Date), i, 1), "mmmm")
        Next

        For i = lYearStart To lYearsAdd
            .CmoYear.AddItem Format(DateSerial(i, 1, 1), "yyyy")
        Next

        .Tag = "Calendar"
        .CmoMonth.ListIndex = Month(Date) - 1
        .CmoYear.ListIndex = Year(Date) - lYearStart
        .Tag = ""
    End With
    Call BuildCal

End Sub

Private Sub CmoMonth_Change()
    If Not Me.Tag = "Calendar" Then BuildCal
End Sub

Private Sub CmoYear_Change()
    If Not Me.Tag = "Calendar" Then BuildCal
End Sub

Sub BuildCal()
    Dim i As Integer
    Dim dTemp As Date
    Dim dTemp2 As Date
    Dim iFirstDay As Integer
    With Me
        .Caption = " " & .CmoMonth.Value & " " & .CmoYear.Value

        dTemp = CDate("01/" & .CmoMonth.Value & "/" & .CmoYear.Value)
        iFirstDay = WeekDay(dTemp, vbSunday)
        .Controls("D" & iFirstDay).SetFocus

        For i = 1 To 42
            With .Controls("D" & i)
                dTemp2 = DateAdd("d", (i - iFirstDay), dTemp)
                .Caption = Format(dTemp2, "d")
                .Tag = dTemp2
                .ControlTipText = Format(dTemp2, "dd/mm/yy")
                'add dates to the buttons
                If Format(dTemp2, "mmmm") = CmoMonth.Value Then
                    If .BackColor <> &H80000016 Then .BackColor = &H80000018
                    If Format(dTemp2, "m/d/yy") = Format(Date, "m/d/yy") Then .SetFocus
                    .Font.Bold = True
                Else
                    If .BackColor <> &H80000016 Then .BackColor = &H8000000F
                    .Font.Bold = False
                End If
                'format the buttons
            End With
        Next
    End With

End Sub
