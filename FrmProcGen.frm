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

Option Explicit
Private Const NEWLINE = 13
Private Const INTAB = 9
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

Public Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    Const CF_UNICODETEXT As Long = 13&
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        GetClipboard = sUniText
    End If
    CloseClipboard
End Function



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
    If OptEntryPoint Then
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & "Dim ErrNo As Integer"
        Txt = Txt & Chr(NEWLINE)
    End If
    
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "Const StrPROCEDURE As String = """ & TxtProcName & "()"" "
    
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    
    Txt = Txt & Chr(INTAB) & "On Error GoTo ErrorHandler"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptEntryPoint Then Txt = Txt & "Restart:"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    
    'Error handling
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & Chr(INTAB) & TxtProcName & " = True"
    If OptEntryPoint Then Txt = Txt & "GracefulExit:"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & "Exit Function"
    If OptEntryPoint Then Txt = Txt & "Exit Sub"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & "ErrorExit:"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(INTAB) & "'***CleanUpCode***"
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & Chr(INTAB) & TxtProcName & " = False" & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    If OptNonEntry Then Txt = Txt & "Exit Function"
    If OptEntryPoint Then Txt = Txt & "Exit Sub"
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & Chr(NEWLINE)
    Txt = Txt & "ErrorHandler:"
    Txt = Txt & Chr(NEWLINE)
    If OptEntryPoint Then
        Txt = Txt & Chr(INTAB) & "If Err.Number >= 1000 And Err.Number <= 1500 Then"
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & Chr(INTAB) & "ErrNo = err.Number"
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & Chr(INTAB) & "CustomErrorHandler (err.Number)"
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & Chr(INTAB) & "If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit"
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & "End If"
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(NEWLINE)
        Txt = Txt & Chr(INTAB) & "If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then"
    End If
    If OptNonEntry Then
        Txt = Txt & Chr(INTAB) & "If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then"
    End If
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
    
    SetClipboard Txt
        Hide
End Sub

Private Sub OptNonEntry_Click()
    TxtRtnVarType = "Boolean"
    OptFunction.Value = True
End Sub

