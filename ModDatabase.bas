Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Private DBPath As String
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

    If DB Is Nothing Then Initialise
    
    Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
    Set SQLQuery = RstResults

Exit Function

ErrorExit:

    Set SQLQuery = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
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
    
    DBPath = ShtSettings.Range("DBPAth")
    
    If DBPath = "" Then
        MsgBox ("No database selected")
        ModDatabase.SelectDB
    End If
    
    Set DB = OpenDatabase(DBPath)

    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
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

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
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
        MsgBox "There was no database selected", vbOKOnly, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    'save database location
    ShtSettings.Range("DBPAth") = FileLoc
    DBPath = ShtSettings.Range("DBPAth")
    
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

