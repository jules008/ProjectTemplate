Attribute VB_Name = "ModSecurity"
'===============================================================
' Module ModSecurity
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModSecurity"

' ===============================================================
' CourseAccessCheck
' Returns whether person is on access list
' ---------------------------------------------------------------
Public Function CourseAccessCheck(CourseNo As String) As Boolean
    Dim StrUsername As String
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    
    Const StrPROCEDURE As String = "CourseAccessCheck()"

    On Error GoTo ErrorHandler

    StrUsername = "'" & Application.Username & "'"
    StrCourseNo = "'" & CourseNo & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                            " CourseNo = " & StrCourseNo & _
                            " AND username = " & StrUsername)
    
    If RstUserList.recordcount = 0 Then
        CourseAccessCheck = False
    Else
        CourseAccessCheck = True
    End If
    
    Set RstUserList = Nothing

    CourseAccessCheck = True

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    CourseAccessCheck = False

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
Private Function RemoveUser(Username As String) As Boolean
    Dim StrUsername As String
    Dim StrCourseNo As String
    Dim CourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset
    
    Const StrPROCEDURE As String = "RemoveUser()"

    On Error GoTo ErrorHandler

    StrUsername = "'" & Username & "'"
    
    If ModDatabase.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    'if courseno is not included, then delete the user from both the user list tables
    'and the course access table
    If CourseNo = "" Then
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM UserList WHERE " & _
                                                "Username = " & StrUsername)
        
        Set RstCourseUserLst = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                                "Username = " & StrUsername)
    Else
    
        'if course no is included, then only delete the user from the course access table
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo & _
                                " AND username = " & StrUsername)
        
    End If
    
    With RstCourseUserLst
        If Not RstCourseUserLst Is Nothing Then
            If .recordcount > 0 Then
                Do While Not .EOF
                    .Delete
                    .MoveNext
                Loop
            End If
        End If
    End With
        
    With RstUserList
        If .recordcount > 0 Then
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
Private Function IsAdmin() As Boolean
    Const StrPROCEDURE As String = "IsAdmin()"

    Dim RstUserList As Recordset
    Dim StrUsername As String
    
    On Error GoTo ErrorHandler

    
    If ModDatabase.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    StrUsername = "'" & Application.Username & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " username = " & StrUsername _
                            & "AND admin = TRUE")
    
    With RstUserList
        If .recordcount > 0 Then
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
Private Function GetAccessList() As Recordset
    Const StrPROCEDURE As String = "GetAccessList()"
    
    Dim StrUsername As String
    Dim StrCourseNo As String
    Dim CourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset

    On Error GoTo ErrorHandler
    
    If CourseNo = "" Then
        Set RstUserList = ModDatabase.SQLQuery("userlist")
    Else
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo)
    End If
    
    If RstUserList.recordcount <> 0 Then
        Set GetAccessList = RstUserList
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
