Attribute VB_Name = "ModAPICalls"
'===============================================================
' Module ModAPICalls
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 26 Jul 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModAPICalls"

' ===============================================================
' ShellExecute
' Executes shell commands
' ---------------------------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

' ===============================================================
' CopyMemory
' Copies blocks of memory from one location to another
' ---------------------------------------------------------------
Public Declare Sub CopyMemory _
Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' ===============================================================
' Sleep
' Pauses execution for a defined number of milliseconds
' ---------------------------------------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

    Dim x  As Long
    Dim y  As Long
   
    x = GetSystemMetrics(SM_CXSCREEN)
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



