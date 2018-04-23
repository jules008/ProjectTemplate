Attribute VB_Name = "ModErrorHandling"
'===============================================================
' Module ModErrorHandling
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModErrorHandling"

' ===============================================================
' CentralErrorHandler
' Handles all system errors
' ---------------------------------------------------------------
Public Function CentralErrorHandler( _
            ByVal ErrModule As String, _
            ByVal ErrProc As String, _
            Optional ByVal ErrFile As String, _
            Optional ByVal EntryPoint As Boolean) As Boolean

    Static ErrMsg As String
    
    Dim iFile As Integer
    Dim ErrNum As Long
    Dim ErrHeader As String
    Dim LogText As String
    
    ErrNum = Err.Number
    ErrMsg = Err.Description
    
    If Len(ErrMsg) = 0 Then ErrMsg = Err.Description
                
    On Error Resume Next
    
    If Len(ErrFile) = 0 Then ErrFile = ThisWorkbook.Name
    
    If Right$(SYS_PATH, 1) <> "\" Then SYS_PATH = SYS_PATH & "\"
    
    ErrHeader = "[" & Application.UserName & "]" & "[" & ErrFile & "]" & ErrModule & "." & ErrProc

    LogText = "  " & ErrHeader & ", Error " & CStr(ErrNum) & ": " & ErrMsg
    
    If Not DEBUG_MODE Then
        
        iFile = FreeFile()
        Open SYS_PATH & FILE_ERROR_LOG For Append As #iFile
        Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); LogText
        If EntryPoint Then Print #iFile,
        Close #iFile
    End If
                
    Debug.Print Format$(Now(), "mm/dd/yy hh:mm:ss"); LogText
    
    If EntryPoint Then
        Debug.Print
        ModLibrary.PerfSettingsOff
        
        If Not DEV_MODE And SEND_ERR_MSG Then SendErrMessage
            SendErrMessage
        ErrMsg = vbNullString
    End If
    
    CentralErrorHandler = DEBUG_MODE
    
    ModLibrary.PerfSettingsOff
End Function

' ===============================================================
' CustomErrorHandler
' Handles system custom errors 1000 - 1500
' ---------------------------------------------------------------
Public Function CustomErrorHandler(ErrorCode As Long, Optional Message As String) As Boolean
    Dim MailSubject As String
    Dim MailBody As String
    
    Const StrPROCEDURE As String = "CustomErrorHandler()"

    On Error Resume Next

    Select Case ErrorCode
        Case UNKNOWN_USER
            
        Case NO_DATABASE_FOUND
            FaultCount1008 = FaultCount1008 + 1
            Debug.Print "Trying to connect to Database....Attempt " & FaultCount1008
            
            If FaultCount1008 <= 3 Then
            
                Application.DisplayStatusBar = True
                Application.StatusBar = "Trying to connect to Database....Attempt " & FaultCount1008
                Application.Wait (Now + TimeValue("0:00:02"))
                Debug.Print FaultCount1008
            Else
                FaultCount1008 = 0
                Application.StatusBar = "System Failed - No Database"
                End
            End If
        
        Case SYSTEM_RESTART
            Debug.Print "system failed - restarting"
            FaultCount1002 = FaultCount1002 + 1

            If FaultCount1002 <= 3 Then
                If Not Initialise Then Err.Raise HANDLED_ERROR
                Application.DisplayStatusBar = True
                Application.StatusBar = "System failed...Restarting Attempt " & FaultCount1002
                Application.Wait (Now + TimeValue("0:00:02"))
            Else
                FaultCount1002 = 0
                Application.StatusBar = "Sysetm Failed"
                End
            End If
            
        Case ACCESS_DENIED
            MsgBox "Sorry you do not have the required Access Level.  " _
                & "Please send a Support Mail if you require access", vbCritical, APP_NAME
        
        Case NO_INI_FILE
            MsgBox "No INI file has been found, so system cannot continue. This can occur if the file " _
                    & "is copied from its location on the T Drive.  Please delete file and create a shortcut instead", vbCritical, APP_NAME
            Application.StatusBar = "System Failed - No INI File"
            End
        
        Case DB_WRONG_VER
            MsgBox "Incorrect Version Database - System cannot continue", vbCritical + vbOKOnly, APP_NAME
            Application.StatusBar = "System Failed - Wrong DB Version"
            End
        
    End Select
    
    Set MailSystem = Nothing

    CustomErrorHandler = True
End Function

' ===============================================================
' SendErrMessage
' Sends an email log file
' ---------------------------------------------------------------
Private Sub SendErrMessage()
    
    On Error Resume Next
    
    If MailSystem Is Nothing Then Set MailSystem = New ClsMailSystem
        
    If Not ModLibrary.OutlookRunning Then
        Shell "Outlook.exe"
    End If

    With MailSystem
        .MailItem.To = "Julian Turner"
        .MailItem.Subject = "Debug Report - " & APP_NAME
        .MailItem.Importance = olImportanceHigh
        .MailItem.Attachments.Add SYS_PATH & FILE_ERROR_LOG
                           .SendEmail
        If SEND_EMAILS Then .SendEmail Else .DisplayEmail
    End With

End Sub
