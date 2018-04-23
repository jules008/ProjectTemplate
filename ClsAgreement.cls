VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsAgreement
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Apr 18
'===============================================================
' Methods
'---------------------------------------------------------------
' DBGet - Gets class from Database
' DBSave - Saves class to Database
' DBDelete(Optional FullDelete As Boolean) - Marks record as deleted or fully deletes
'===============================================================

Option Explicit
Private pDeleted As Date
Private pCrewNo As String
Private pTimeSlots As String
Private pCrewMember As ClsCrewMember
Private pAADate As Date
Private pContractType As String
Private pHrsPW As Integer
Private pNoWeeks As Integer
Private pStation As String
Private pRevDateDue As Date

'---------------------------------------------------------------
Public Property Get Deleted() As Date
    Deleted = pDeleted
End Property

Public Property Let Deleted(ByVal vNewValue As Date)
    pDeleted = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewNo
End Property

Public Property Let CrewNo(ByVal vNewValue As String)
    pCrewNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TimeSlots() As String
    TimeSlots = pTimeSlots
End Property

Public Property Let TimeSlots(ByVal vNewValue As String)
    pTimeSlots = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CrewMember() As ClsCrewMember
    Set CrewMember = pCrewMember
End Property

Public Property Let CrewMember(ByVal vNewValue As ClsCrewMember)
    Set pCrewMember = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AADate() As Date
    AADate = pAADate
End Property

Public Property Let AADate(ByVal vNewValue As Date)
    pAADate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ContractType() As String
    ContractType = pContractType
End Property

Public Property Let ContractType(ByVal vNewValue As String)
    pContractType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get HrsPW() As Integer
    HrsPW = pHrsPW
End Property

Public Property Let HrsPW(ByVal vNewValue As Integer)
    pHrsPW = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoWeeks() As Integer
    NoWeeks = pNoWeeks
End Property

Public Property Let NoWeeks(ByVal vNewValue As Integer)
    pNoWeeks = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Station() As String
    Station = pStation
End Property

Public Property Let Station(ByVal vNewValue As String)
    pStation = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RevDateDue() As Date
    RevDateDue = pRevDateDue
End Property

Public Property Let RevDateDue(ByVal vNewValue As Date)
    pRevDateDue = vNewValue
End Property

'---------------------------------------------------------------
' ===============================================================
' Method DBGet
' Gets class from Database
'---------------------------------------------------------------
Public Sub DBGet()
    Dim RstAgreement As Recordset

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    With RstAgreement
        If .RecordCount > 0 Then
            If Not IsNull(!Deleted) Then pDeleted = !Deleted
            If Not IsNull(!CrewNo) Then pCrewNo = !CrewNo
            If Not IsNull(!TimeSlots) Then pTimeSlots = !TimeSlots
            If Not IsNull(!CrewMember) Then pCrewMember = !CrewMember
            If Not IsNull(!AADate) Then pAADate = !AADate
            If Not IsNull(!ContractType) Then pContractType = !ContractType
            If Not IsNull(!HrsPW) Then pHrsPW = !HrsPW
            If Not IsNull(!NoWeeks) Then pNoWeeks = !NoWeeks
            If Not IsNull(!Station) Then pStation = !Station
            If Not IsNull(!RevDateDue) Then pRevDateDue = !RevDateDue

            pCrewMember.DBGet
        End If
    End With
Set RstAgreement = Nothing
End Sub

' ===============================================================
' Method DBSave
' Saves class to Database
'---------------------------------------------------------------
Public Sub DBSave()
    Dim RstAgreement As Recordset
    Dim RstMaxNo As Recordset
    Dim LastNo As Integer

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    Set RstMaxNo = ModDatabase.SQLQuery("SELECT MAX(CrewNo) FROM TblAgreement ")

    If RstMaxNo.Fields(0).Value <> 0 Then
        LastNo = RstMaxNo.Fields(0).Value
    Else
        LastNo = 0
    End If

    With RstAgreement
        If .RecordCount = 0 Then
            .addnew
            pCrewNo = LastNo + 1
        Else
            .Edit
        End If
        !CrewNo = pCrewNo
        !TimeSlots = pTimeSlots
        !CrewMember = pCrewMember
        !AADate = pAADate
        !ContractType = pContractType
        !HrsPW = pHrsPW
        !NoWeeks = pNoWeeks
        !Station = pStation
        !RevDateDue = pRevDateDue
        .Update

    End With
    Set RstAgreement = Nothing
    Set RstMaxNo = Nothing
End Sub

' ===============================================================
' Method DBDelete(Optional FullDelete As Boolean)
' Marks record as deleted or fully deletes
'---------------------------------------------------------------
Public Sub DBDelete(Optional FullDelete As Boolean)
    Dim RstAgreement As Recordset
    Dim i As Integer

    Set RstAgreement = ModDatabase.SQLQuery("SELECT * FROM TblAgreement WHERE CrewNo = " & pCrewNo & " AND Deleted IS NULL")
    With RstAgreement
        For i = .RecordCount To 1 Step -1
            If FullDelete Then
                .Delete
                .MoveNext
            Else
                .Edit
                !Deleted = Now
                .Update
            End If
        Next
    End With

    Set RstAgreement = Nothing
End Sub


' ===============================================================
Private Sub Class_Initialize()
    Set pCrewMember = New ClsCrewMember

    pCrewMember.SetParent ObjPtr(Me)
    pAADate.SetParent ObjPtr(Me)
    pContractType.SetParent ObjPtr(Me)
    pHrsPW.SetParent ObjPtr(Me)
    pNoWeeks.SetParent ObjPtr(Me)
    pStation.SetParent ObjPtr(Me)
    pRevDateDue.SetParent ObjPtr(Me)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pCrewMember = Nothing

    pDeleted.SetParent 0
    pCrewNo.SetParent 0
    pTimeSlots.SetParent 0
    pCrewMember.SetParent 0
    pAADate.SetParent 0
    pContractType.SetParent 0
    pHrsPW.SetParent 0
    pNoWeeks.SetParent 0
    pStation.SetParent 0
    pRevDateDue.SetParent 0
End Sub

'---------------------------------------------------------------
