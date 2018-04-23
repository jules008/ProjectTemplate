Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModCloseDown"

' ===============================================================
' Terminate
' Closedown processing
' ---------------------------------------------------------------
Public Function Terminate() As Boolean
    Const StrPROCEDURE As String = "Terminate()"

    On Error GoTo ErrorHandler

    ModDatabase.DBTerminate
    
    Terminate = True

Exit Function

ErrorExit:

    ModDatabase.DBTerminate

    
    Terminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

