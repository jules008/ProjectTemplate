VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsMailItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsMailItem
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 12 Oct 16
'===============================================================
Option Explicit
Private pEmailNo As Integer
Private pCrewNo As String
Private pMailTo As String
Private pMailFrom As String
Private pCC As String
Private pSubject As String
Private pBody As String
Private ptimeSent As Date

'----------------------------------------------------------
Public Property Get EmailNo() As Integer
    EmailNo = pEmailNo
End Property

Public Property Let EmailNo(ByVal vNewValue As Integer)
    pEmailNo = vNewValue
End Property

'----------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewNo
End Property

Public Property Let CrewNo(ByVal vNewValue As String)
    pCrewNo = vNewValue
End Property

'----------------------------------------------------------
Public Property Get MailTo() As String
    MailTo = pMailTo
End Property

Public Property Let MailTo(ByVal vNewValue As String)
    pMailTo = vNewValue
End Property

'----------------------------------------------------------
Public Property Get MailFrom() As String
    MailFrom = pMailFrom
End Property

Public Property Let MailFrom(ByVal vNewValue As String)
    pMailFrom = vNewValue
End Property

'----------------------------------------------------------
Public Property Get CC() As String
    CC = pCC
End Property

Public Property Let CC(ByVal vNewValue As String)
    pCC = vNewValue
End Property

'----------------------------------------------------------
Public Property Get Subject() As String
    Subject = pSubject
End Property

Public Property Let Subject(ByVal vNewValue As String)
    pSubject = vNewValue
End Property

'----------------------------------------------------------
Public Property Get Body() As String
    Body = pBody
End Property

Public Property Let Body(ByVal vNewValue As String)
    pBody = vNewValue
End Property

'----------------------------------------------------------
Public Property Get timeSent() As Date
    timeSent = ptimeSent
End Property

Public Property Let timeSent(ByVal vNewValue As Date)
    ptimeSent = vNewValue
End Property

'===============================================================
'Method LoadDB
'load Item from database
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstMailItems As Recordset
    
    Set RstMailItems = ModDatabase.SQLQuery("SELECT * FROM mailitems WHERE EmailNo = " & pEmailNo & _
                                        " AND deleted IS NULL")
                                        
    With RstMailItems
        If .RecordCount > 0 Then
            If Not IsNull(!EmailNo) Then pEmailNo = !EmailNo
            If Not IsNull(!CrewNo) Then pCrewNo = !CrewNo
            If Not IsNull(!MailTo) Then pMailTo = !MailTo
            If Not IsNull(!MailFrom) Then pMailFrom = !MailFrom
            If Not IsNull(!Subject) Then pSubject = !Subject
            If Not IsNull(!Body) Then pBody = !Body
            If Not IsNull(!timeSent) Then ptimeSent = !timeSent
        End If
    End With
    Set RstMailItems = Nothing
End Sub

'===============================================================
'Method NewDB
'Creates a new item in the DB
'---------------------------------------------------------------
Public Function NewDB() As Integer
    Dim RstMailItems As Recordset
    
    Set RstMailItems = ModDatabase.SQLQuery("mailitems")
        
    With RstMailItems
        .AddNew
        NewDB = !EmailNo
        !CrewNo = pCrewNo
        .Update
    End With
    
    Set RstMailItems = Nothing
End Function

'===============================================================
'Method UpdateDB
'Update session to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstMailItems As Recordset

    Set RstMailItems = ModDatabase.SQLQuery("SELECT * FROM mailitems WHERE EmailNo = " & pEmailNo & _
                                        " AND deleted IS NULL")
    
    
    With RstMailItems
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !MailTo = pMailTo
            !MailFrom = pMailFrom
            !Subject = pSubject
            !Body = pBody
            !timeSent = ptimeSent
            .Update
            UpdateDB = True
        End If
    End With
    
    Set RstMailItems = Nothing

End Function

'===============================================================
'Method DeleteDB
' Delete session from database
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstMailItems As Recordset
    
    Set RstMailItems = ModDatabase.SQLQuery("SELECT * FROM mailitems WHERE EmailNo = " & pEmailNo & _
                                        " AND deleted IS NULL")
        
    With RstMailItems
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
        End If
    End With
    Set RstMailItems = Nothing
End Sub
