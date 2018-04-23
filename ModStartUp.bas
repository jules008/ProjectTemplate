Attribute VB_Name = "ModStartUp"
'===============================================================
' Module ModStartUp
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
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

    SYS_PATH = ThisWorkbook.Path & INI_FILE_PATH

    Application.StatusBar = "Reading INI File....."
    
    If Not ReadINIFile Then Err.Raise HANDLED_ERROR
    
    Application.StatusBar = "Connecting to DB....."
    
    If Not ModDatabase.DBConnect Then Err.Raise HANDLED_ERROR
    
    Application.StatusBar = "Checking DB Version....."
    
    If ModDatabase.GetDBVer <> DB_VER Then Err.Raise DB_WRONG_VER
    
    Application.StatusBar = "Finding User....."
    'get username of current user
    If Not ModStartUp.GetUserName Then Err.Raise HANDLED_ERROR
    
    'Show any messages
    If Not MessageCheck Then Err.Raise HANDLED_ERROR
    


    Initialise = True

Exit Function

ErrorExit:

    Initialise = False
    
Exit Function

ErrorHandler:
        
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume Next
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' GetUserName
' gets username from windows, or test user if in test mode
' ---------------------------------------------------------------
Public Function GetUserName() As Boolean
    Dim UserName As String
    Dim CharPos As Integer
    
    Const StrPROCEDURE As String = "GetUserName()"

    On Error GoTo ErrorHandler
    
    If Not UpdateUsername Then Err.Raise HANDLED_ERROR
    
    If DEV_MODE Then
'       If ShtSettings.Range("C15") = True Then
'            UserName = ShtSettings.Range("Test_User")
'        Else
        UserName = "Julian Turner"
'        End If
    Else
        UserName = Application.UserName
    End If
    
    If UserName = "" Then Err.Raise HANDLED_ERROR, , "No Username"

    UserName = Replace(UserName, "'", "")
    
GracefulExit:
    
    GetUserName = True

Exit Function

ErrorExit:

    GetUserName = False

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

' ===============================================================
' ReadINIFile
' Gets start up variables from ini file
' ---------------------------------------------------------------
Private Function ReadINIFile() As Boolean
    Dim DebugMode As String
    Dim EnablePrint As String
    Dim DBPath As String
    Dim SendEmails As String
    Dim DevMode As String
    Dim INIFile As Integer
    
    Const StrPROCEDURE As String = "ReadINIFile()"

    On Error GoTo ErrorHandler
       
    INIFile = FreeFile()
       
    Debug.Print SYS_PATH & INI_FILE_NAME
    
    If Dir(SYS_PATH & INI_FILE_NAME) = "" Then Err.Raise NO_INI_FILE
    
    Open SYS_PATH & INI_FILE_NAME For Input As #INIFile
    
    Line Input #INIFile, DebugMode
    Line Input #INIFile, SendEmails
    Line Input #INIFile, EnablePrint
    Line Input #INIFile, DBPath
    Line Input #INIFile, DevMode
    
    Close #INIFile
    
    DEBUG_MODE = CBool(DebugMode)
    SEND_EMAILS = CBool(SendEmails)
    ENABLE_PRINT = CBool(EnablePrint)
    DB_PATH = DBPath
    DEV_MODE = CBool(DevMode)
    
    If STOP_FLAG = True Then Stop
    
    If MAINT_MSG <> "" Then
        MsgBox MAINT_MSG, vbExclamation, APP_NAME
        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
        
    End If
    
    
GracefulExit:
    
    ReadINIFile = True
    Application.DisplayAlerts = True

Exit Function

ErrorExit:

    ReadINIFile = False
    Application.DisplayAlerts = True

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
        Resume ErrorExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' MessageCheck
' Checks to see if the user message has been read
' ---------------------------------------------------------------
Public Function MessageCheck() As Boolean
    Dim StrMessage As String
    Dim RstMessage As Recordset
    
    Const StrPROCEDURE As String = "MessageCheck()"

    On Error GoTo ErrorHandler
    
'    If CurrentUser.AccessLvl >= BasicLvl_1 Then
'        If Not CurrentUser.MessageRead Then
'
'            Set RstMessage = SQLQuery("TblMessage")
'
'            If RstMessage.RecordCount > 0 Then StrMessage = RstMessage.Fields(0)
'            MsgBox StrMessage, vbOKOnly + vbInformation, APP_NAME
'            CurrentUser.MessageRead = True
'            CurrentUser.DBSave
'
'        End If
'    End If
    
    Set RstMessage = Nothing
    
    MessageCheck = True

Exit Function

ErrorExit:
    Set RstMessage = Nothing
'    ***CleanUpCode***
    MessageCheck = False

Exit Function

ErrorHandler:
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateUsername
' Checks to see whether username needs to be changed and then updates
' ---------------------------------------------------------------
Private Function UpdateUsername() As Boolean
    Const StrPROCEDURE As String = "UpdateUsername()"

    On Error GoTo ErrorHandler

    If Application.UserName = "PaulJ Wright" Then Application.UserName = "Paul Wright"
    
    If Application.UserName = "Ian Taylor (LRF)" Then Application.UserName = "Ian Taylor (LFR)"

    If Application.UserName = "#" Then Application.UserName = "Samuel Hayward"
    
    UpdateUsername = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    UpdateUsername = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
