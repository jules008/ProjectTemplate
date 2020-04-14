Attribute VB_Name = "ModSecurity"
'===============================================================
' Module ModSecurity
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModSecurity"

' ===============================================================
' LogUserOn
' Logs on user and assigns access level.  Terminates if user is not known
' ---------------------------------------------------------------
Public Function LogUserOn(UserName As String) As Boolean
    Const StrPROCEDURE As String = "LogUserOn()"

    On Error GoTo ErrorHandler

    If UserName = "" Then Err.Raise HANDLED_ERROR, , "Username blank"
    
    CurrentUser.DBGet UserName
    
    Debug.Print CurrentUser.UserName & " Logged on"
    
GracefulExit:

    LogUserOn = True

Exit Function

ErrorExit:

    '***CleanUpCode***
    LogUserOn = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

