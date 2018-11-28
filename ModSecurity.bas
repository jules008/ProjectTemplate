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
' StationAccessCheck
' Returns whether person is on access list
' ---------------------------------------------------------------
Public Function StationAccessCheck(Station1 As Integer, Station2 As Integer) As EnumTriState
    Dim StrUserName As String
    Dim StrCourseNo As String
    Dim Stations() As String
    Dim RstUserList As Recordset
    
    Const StrPROCEDURE As String = "StationAccessCheck()"

    On Error GoTo ErrorHandler

    Set RstUserList = ModDatabase.SQLQuery("SELECT Stations FROM TblPerson WHERE " & _
                            " username = '" & CurrentUser.UserName & "'")
    
    StationAccessCheck = xFalse
    
    With RstUserList
        If .RecordCount > 0 Then
            Stations = Split(!Stations, ";")
            If Stations(Station1 - 1) = 1 Then StationAccessCheck = xTrue
            
            If Station2 > 0 Then
                If Stations(Station2 - 1) = 1 Then StationAccessCheck = xTrue
            End If
        End If
    End With
    
    Set RstUserList = Nothing

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    StationAccessCheck = xError

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RemoveUser
' Removes user from access list for course
' ---------------------------------------------------------------
Public Function RemoveUser(UserName As String) As Boolean
    Dim StrUserName As String
    Dim StrCourseNo As String
    Dim CourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset
    
    Const StrPROCEDURE As String = "RemoveUser()"

    On Error GoTo ErrorHandler
    
    If DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE " & _
                                            "Username = '" & UserName & "'")
        
    With RstUserList
        If .RecordCount > 0 Then
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End If
    End With
    
    
    Set RstUserList = Nothing
    Set RstCourseUserLst = Nothing

    RemoveUser = True

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    Set RstCourseUserLst = Nothing
    RemoveUser = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' IsAdmin
' Checks whether person is an admin
' ---------------------------------------------------------------
Public Function IsAdmin() As Boolean
    Const StrPROCEDURE As String = "IsAdmin()"

    Dim RstUserList As Recordset
    Dim StrUserName As String
    
    On Error GoTo ErrorHandler

    
    If ModDatabase.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    StrUserName = "'" & Application.UserName & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " username = " & StrUserName _
                            & "AND admin = TRUE")
    
    With RstUserList
        If .RecordCount > 0 Then
            IsAdmin = True
        Else
            IsAdmin = False
        End If
    End With
    
    Set RstUserList = Nothing

    IsAdmin = True

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    IsAdmin = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' GetAccessList
' Returns access list for course
' ---------------------------------------------------------------
Public Function GetAccessList() As Recordset
    Dim RstUserList As Recordset
    
    Const StrPROCEDURE As String = "GetAccessList()"
    
    On Error GoTo ErrorHandler
    
    Set RstUserList = ModDatabase.SQLQuery("TblPerson")
    
    If RstUserList.RecordCount > 0 Then
        Set GetAccessList = RstUserList
    Else
        Set GetAccessList = Nothing
    End If
    
    Set RstUserList = Nothing

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    Set GetAccessList = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' AddUpdateUser
' Adds or updates user
' ---------------------------------------------------------------
Public Function AddUpdateUser(User As ClsPerson) As Boolean
    Dim RstUserList As Recordset
    
    Const StrPROCEDURE As String = "AddUpdateUser()"

    On Error GoTo ErrorHandler

    If User.UserName = "" Then
        User.UserName = User.Forename & " " & User.Surname
    End If
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE " & _
                                        "username = '" & User.UserName & "'")
    With RstUserList
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        !CrewNo = User.CrewNo
        !RankGrade = User.RankGrade
        !Forename = User.Forename
        !Surname = User.Surname
        !Role = User.Role
        !Stations = User.Stations
        !UserName = User.UserName
        .Update
    
    End With
    
    Set RstUserList = Nothing
    AddUpdateUser = True
Exit Function

ErrorExit:
    Set RstUserList = Nothing
    AddUpdateUser = False

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
' GetUserDetails
' Returns user details in Recordset
' ---------------------------------------------------------------
Public Function GetUserDetails(UserName As String) As Recordset
    Dim RstUserList As Recordset
    
    On Error Resume Next
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM TblPerson WHERE " & _
                            " UserName = '" & UserName & "'")
                            
    If RstUserList.RecordCount > 0 Then
        Set GetUserDetails = RstUserList
    Else
        Set GetUserDetails = Nothing
    End If
    
    Set RstUserList = Nothing
    
End Function
