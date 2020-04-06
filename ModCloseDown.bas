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
    
    If Not EndGlobalClasses Then Err.Raise HANDLED_ERROR

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

' ===============================================================
' EndGlobalClasses
' initialises or terminates all global classes
' ---------------------------------------------------------------
Private Function EndGlobalClasses() As Boolean
    Const StrPROCEDURE As String = "EndGlobalClasses()"

    On Error GoTo ErrorHandler

    Set CurrentUser = Nothing
    
    EndGlobalClasses = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    EndGlobalClasses = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

