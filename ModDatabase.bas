Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef

' ===============================================================
' SQLQuery
' Queries database with given SQL script
' ---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "SQLQuery()"

    On Error GoTo ErrorHandler
      
Restart:
    Application.StatusBar = ""

    If DB Is Nothing Then
        Err.Raise NO_DATABASE_FOUND, Description:="Unable to connect to database"
    Else
        If FaultCount1008 > 0 Then FaultCount1008 = 0
    
        Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
        Set SQLQuery = RstResults
    End If
    
    Set RstResults = Nothing
    
Exit Function

ErrorExit:

    Set RstResults = Nothing

    Set SQLQuery = Nothing
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            Resume Restart
        Else
            Err.Raise HANDLED_ERROR
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DBConnect
' Provides path to database
' ---------------------------------------------------------------
Public Function DBConnect() As Boolean
    Const StrPROCEDURE As String = "DBConnect()"

    On Error GoTo ErrorHandler

    Set DB = OpenDatabase(DB_PATH & DB_FILE_NAME)
  
    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

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
' DBTerminate
' Disconnects and closes down DB connection
' ---------------------------------------------------------------
Public Function DBTerminate() As Boolean
    Const StrPROCEDURE As String = "DBTerminate()"

    On Error GoTo ErrorHandler

    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing

    DBTerminate = True

Exit Function

ErrorExit:

    DBTerminate = False

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
' SelectDB
' Selects DB to connect to
' ---------------------------------------------------------------
Public Function SelectDB() As Boolean
    Const StrPROCEDURE As String = "SelectDB()"

    On Error GoTo ErrorHandler
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly + vbExclamation, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    DB_PATH = FileLoc
    
    Set DlgOpen = Nothing

    SelectDB = True

Exit Function

ErrorExit:

    Set DlgOpen = Nothing
    SelectDB = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScript()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
    Dim Binary As String
    
    Dim Fld As DAO.Field
    
    DBConnect
        
    Set RstTable = SQLQuery("TblDBVersion")
    
    'check preceding DB Version
    If RstTable.Fields(0) <> "v1,393" Then
        MsgBox "Database needs to be upgraded to v1,393 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    MsgBox "Import tables TblVehicle and TblVehicleType"
    
    'Table changes
    
    ' delete old vehicle table and add new
    DB.Execute "SELECT * INTO TblVehicleOLD FROM TblVehicle"
    DB.Execute "DROP TABLE TblVehicle"
    DB.Execute "SELECT * INTO TblVehicle FROM TblVehicleNEW"
    DB.Execute "DROP TABLE TblVehicleNEW"
    
    ' delete old vehicleType table and add new
    DB.Execute "SELECT * INTO TblVehicleTypeOLD FROM TblVehicleType"
    DB.Execute "DROP TABLE TblVehicleType"
    DB.Execute "SELECT * INTO TblVehicleType FROM TblVehicleTypeNEW"
    DB.Execute "DROP TABLE TblVehicleTypeNEW"
    
    'clear new issue flag
    
    DB.Execute "SELECT * INTO TblAssetOLD FROM TblAsset"
    Set RstTable = SQLQuery("TblAsset")
    
    i = 1
    With RstTable
        Do While Not .EOF
            Debug.Print !AssetNo
            Binary = !AllowedOrderReasons
            
            If Len(Binary) <> 13 Then
                Binary = Left(Binary, 13)
                Debug.Print "Length corrected on Asset " & !AssetNo
            End If
            
            Binary = Left(Binary, 12) & "0"
            
            .Edit
            !AllowedOrderReasons = Binary
            .Update
            .MoveNext
            i = i + 1
        Loop
    
    End With
    
    'update DB Version
    Set RstTable = SQLQuery("TblDBVersion")
    
    With RstTable
        .Edit
        .Fields(0) = "v1,394"
        .Update
    End With
    
    UpdateSysMsg
    
    MsgBox "Database successfully updated", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing
    
End Sub
              
' ===============================================================
' UpdateDBScriptUndo
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScriptUndo()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
        
    Dim Fld As DAO.Field
        
    DBConnect
    
    Set RstTable = SQLQuery("TblDBVersion")

    If RstTable.Fields(0) <> "v1,394" Then
        MsgBox "Database needs to be upgraded to v1,394 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
       
    'Undo Vehicle update
    DB.Execute "SELECT * INTO TblVehicleNEW FROM TblVehicle"
    DB.Execute "DROP TABLE TblVehicle"
    DB.Execute "SELECT * INTO TblVehicle FROM TblVehicleOLD"
    DB.Execute "DROP TABLE TblVehicleOLD"
    
    DB.Execute "SELECT * INTO TblVehicleTypeNEW FROM TblVehicleType"
    DB.Execute "DROP TABLE TblVehicleType"
    DB.Execute "SELECT * INTO TblVehicleType FROM TblVehicleTypeOLD"
    DB.Execute "DROP TABLE TblVehicleTypeOLD"
     
    Set RstTable = SQLQuery("TblDBVersion")

    'undo new issue update
    DB.Execute "DROP TABLE TblAsset"
    DB.Execute "SELECT * INTO TblAsset FROM TblAssetOLD"
    DB.Execute "DROP TABLE TblAssetOLD"
    
    With RstTable
        .Edit
        .Fields(0) = "v1,393"
        .Update
    End With
    
    MsgBox "Database reset successfully", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing

End Sub

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetDBVer = DBVer.Fields(0)

    Debug.Print DBVer.Fields(0)
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateSysMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub UpdateSysMsg()
    Dim RstMessage As Recordset
    
    Set RstMessage = SQLQuery("TblMessage")
    
    With RstMessage
        If .RecordCount = 0 Then
            .addnew
        Else
            .Edit
        End If
        
        .Fields("SystemMessage") = "Version " & VERSION & " - What's New" _
                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Bug Fix - Hidden Assets" _
                    & Chr(13) & ""
        
        .Fields("ReleaseNotes") = "Software Version: " & VERSION _
                    & Chr(13) & "Database Version: " & DB_VER _
                    & Chr(13) & "Date: " & VER_DATE _
                    & Chr(13) & "" _
                    & Chr(13) & "- Bug Fix - Hidden Assets - Had ANOTHER go at fixing the hidden assets bug.  Hopefully fixed now" _
                    & Chr(13) & ""
        .Update
    End With
    
    'reset read flags
    DB.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
    
    Set RstMessage = Nothing

End Sub

' ===============================================================
' ShowUsers
' Show users logged onto system
' ---------------------------------------------------------------
Public Sub ShowUsers()
    Dim RstUsers As Recordset
    
    Set RstUsers = SQLQuery("TblUsers")
    
    With RstUsers
        Debug.Print
        Do While Not .EOF
            Debug.Print "User: " & .Fields(0) & " - Logged on: " & .Fields(1)
            .MoveNext
        Loop
    End With
    
    Set RstUsers = Nothing
End Sub
