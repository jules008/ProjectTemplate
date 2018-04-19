Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModStartUp"

' ===============================================================
' Initialise
' Creates the environment for system start up
' ---------------------------------------------------------------
Public Function Initialise() As Boolean
    Const StrPROCEDURE As String = "Initialise()"

    On Error GoTo ErrorHandler

    Terminate
    Set MailSystem = New ClsMailSystem
    
    If Not ModDatabase.DBConnect Then Err.Raise HANDLED_ERROR


    Initialise = True

Exit Function

ErrorExit:

    Set MailSystem = Nothing
    Initialise = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
