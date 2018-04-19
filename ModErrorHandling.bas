Attribute VB_Name = "ModErrorHandling"
'===============================================================
' Module ModErrorHandling
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModErrorHandling"

' ===============================================================
' CentralErrorHandler
' Handles all system errors
' ---------------------------------------------------------------
Public Function CentralErrorHandler( _
            ByVal sModule As String, _
            ByVal sProc As String, _
            Optional ByVal sFile As String, _
            Optional ByVal bEntryPoint As Boolean) As Boolean

    Static sErrMsg As String
    
    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim sPath As String
    Dim sLogText As String
    Dim ErrMsgTxt As String
    
    ' Grab the error info before it's cleared by
    ' On Error Resume Next below.
    lErrNum = Err.Number
    ' If this is a user cancel, set the silent error flag
    ' message. This will cause the error to be ignored.
'    If lErrNum = USER_CANCEL Then sErrMsg = SILENT_ERROR
    ' If this is the originating error, the static error
    ' message variable will be empty. In that case, store
    ' the originating error message in the static variable.
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description
                

    ' We cannot allow errors in the central error handler.
    On Error Resume Next
    
    ' Load the default filename if required.
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name
    
    ' Get the application directory.
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Create the error text to be logged.
    ErrMsgTxt = "Sorry, there has been an error.  An Error Log File has been created.  Would " _
                & " like to email this for further investigation?"
        
    sLogText = "  " & sFullSource & ", Error " & _
                        CStr(lErrNum) & ": " & sErrMsg
    
    ' Open the log file, write out the error information and
    ' close the log file.
    If OUTPUT_MODE = "Log" Then
        Dim Response As Integer
        
        iFile = FreeFile()
        Open sPath & FILE_ERROR_LOG For Append As #iFile
        Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
        If bEntryPoint Then Print #iFile,
        Close #iFile
                
    Else
        Debug.Print Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
        If bEntryPoint Then Debug.Print
    End If
    
    ' Do not display or debug silent errors.
'    If sErrMsg <> SILENT_ERROR Then

    ' Show the error message when we reach the entry point
    ' procedure or immediately if we are in debug mode.
    If bEntryPoint Or DEBUG_MODE Then
        Application.ScreenUpdating = True
        Response = MsgBox(ErrMsgTxt, vbYesNo, APP_NAME)
    
        If Response = 6 Then
            With MailSystem
                .MailItem.To = "Julian Turner"
                .MailItem.Subject = "Debug Report - " & APP_NAME
                .MailItem.Importance = olImportanceHigh
                .MailItem.Attachments.Add sPath & FILE_ERROR_LOG
                .MailItem.Body = "Please add any further information such " _
                                   & "what you were doing at the time of the error" _
                                   & ", and what candidate were you working on etc "
                .DisplayEmail
            End With
        End If
        
        ' Clear the static error message variable once
        ' we've reached the entry point so that we're ready
        ' to handle the next error.
        sErrMsg = vbNullString
    End If
    
    ' The return vale is the debug mode status.
    CentralErrorHandler = DEBUG_MODE
    
'    Else
'        ' If this is a silent error, clear the static error
'        ' message variable when we reach the entry point.
'        If bEntryPoint Then sErrMsg = vbNullString
'        CentralErrorHandler = False
'    End If
    
End Function


