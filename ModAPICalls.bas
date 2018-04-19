Attribute VB_Name = "ModAPICalls"
'===============================================================
' Module ModAPICalls
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAPICalls"

' ===============================================================
' CopyMemory
' Copies blocks of memory from one location to another
' ---------------------------------------------------------------
Public Declare Sub CopyMemory _
Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' ===============================================================
' GetScreenHeight
' Gets the screen height from the API
' ---------------------------------------------------------------
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1

Public Function GetScreenHeight() As Integer
    Const StrPROCEDURE As String = "GetScreenHeight()"

    On Error GoTo ErrorHandler

    Dim X  As Long
    Dim y  As Long
   
    X = GetSystemMetrics(SM_CXSCREEN)
    y = GetSystemMetrics(SM_CYSCREEN)

    GetScreenHeight = y

    GetScreenHeight = True

Exit Function

ErrorExit:

    GetScreenHeight = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



