VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProcGen 
   Caption         =   "Procedure Generator"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8055
   OleObjectBlob   =   "FrmProcGen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProcGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Form FrmProcGen
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================

Option Explicit
Private Const NEWLINE = 13
Private Const INTAB = 9


Private Sub BtnAddInVar_Click()
    With LstInputVars
        .AddItem
        .List(.ListCount - 1, 0) = TxtInVarName
        .List(.ListCount - 1, 1) = TxtInVarType
        
    End With
    
    TxtInVarName = ""
    TxtInVarType = ""

End Sub

Private Sub BtnPaste_Click()
    
    Dim obj As New DataObject
    Dim Txt As String
    Dim i As Integer
    
    'header
    
    Txt = Txt & Chr(NEWLINE) & "' ==============================================================="
    Txt = Txt & Chr(NEWLINE) & "' " & TxtProcName
    Txt = Txt & Chr(NEWLINE) & "' " & TxtPurpose
    Txt = Txt & Chr(NEWLINE) & "' ---------------------------------------------------------------"
    Txt = Txt & Chr(NEWLINE)

    
    If OptPrivate Then Txt = Txt & "Private "
    If OptPublic Then Txt = Txt & "Public "
    If OptSub Then Txt = Txt & "Sub "
    If OptFunction Then Txt = Txt & "Function "
    
    Txt = Txt & TxtProcName & "("
    
    With LstInputVars
        For i = 0 To .ListCount - 1
            Txt = Txt & .List(i, 0) & " As " & .List(i, 1)
            If i <> .ListCount - 1 Then Txt = Txt & ", "
        Next
    End With
    
    Txt = Txt & ")"
    
    If OptFunction Then Txt = Txt & " As " & TxtRtnVarType

    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "Const StrPROCEDURE As String = """ & TxtProcName & "()"" "
    
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    
    Txt = Txt & Chr(INTAB) & "On Error GoTo ErrorHandler"
    
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    
    'Error handling
    If OptNonEntry Then Txt = Txt & Chr(INTAB) & TxtProcName & " = True"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & "Exit Function"
    If OptEntryPoint Then Txt = Txt & "Exit Sub"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & "ErrorExit:"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "***CleanUpCode***"
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & Chr(INTAB) & TxtProcName & " = False" & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & "Exit Function"
    If OptEntryPoint Then Txt = Txt & "Exit Sub"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & "ErrorHandler:"
    If OptEntryPoint Then Txt = Txt & Chr(INTAB) & "If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then"
    If OptNonEntry Then Txt = Txt & Chr(INTAB) & "If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & Chr(INTAB) & "Stop"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & Chr(INTAB) & "Resume"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "Else"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & Chr(INTAB) & "Resume ErrorExit"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "End If"
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & "End Function"
    If OptEntryPoint Then Txt = Txt & "End Sub"
    
    obj.SetText Txt
    obj.PutInClipboard
    Hide
End Sub

Private Sub OptNonEntry_Click()
    TxtRtnVarType = "Boolean"
    OptFunction.Value = True
End Sub

Private Sub UserForm_Click()

End Sub
