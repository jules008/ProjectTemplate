VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUILineitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class Lineitems
' v0,0 - Initial Version
' v0,1 - Bug fix to Width Property
' v0,2 - Add option to disable OnAction for Lineitem
' v0,3 - OnAction now passed as a parameter
'---------------------------------------------------------------
' Date - 12 May 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' ReOrder - re-orders the shapes so that the Lineitems sit on top of the screen
' Text - Adds text to the line item
'===============================================================

Option Explicit
Private pNoColumns As Integer
Private pColumns() As Integer
Private PUILineitems As Collection
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pStyle As TypeStyle
Private pRowOffset As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get NoColumns() As Integer
    NoColumns = pNoColumns
End Property

Public Property Let NoColumns(ByVal vNewValue As Integer)
    pNoColumns = vNewValue
    ReDim pColumns(0 To NoColumns - 1)
End Property

'---------------------------------------------------------------
Public Property Get ColumnWidth(i As Integer) As Integer
    ColumnWidth = pColumns(i)
End Property

Public Property Let Columns(vNewValue As String)
    Dim i As Integer
    Dim ColArray() As String
    
    ColArray() = Split(vNewValue, ":")
    
    For i = 0 To pNoColumns - 1
        pColumns(i) = CInt(ColArray(i))
    Next
End Property

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop - Parent.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue + Parent.Top
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue + Parent.Left
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pHeight
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pWidth
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get RowOffset() As Integer
    RowOffset = pRowOffset
End Property

Public Property Let RowOffset(ByVal vNewValue As Integer)
    pRowOffset = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIFrame
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = PUILineitems.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUILineitem
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = PUILineitems.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UILineitem As ClsUILineitem)
    UILineitem.SetParent ObjPtr(Me)
    PUILineitems.Add UILineitem, Key:=CStr(UILineitem.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    PUILineitems.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = PUILineitems.Count
End Function

' ===============================================================
' Method Text
' Adds text to the line item
'---------------------------------------------------------------
Public Sub Text(Row As Integer, Col As Integer, ItemText As String, Optional OnAction As String)
    Dim DashObj As ClsUIDashObj
    Dim UILineitem As ClsUILineitem
    Dim LeftOffset As Integer
    Dim i As Integer
    
    Do While i <> Col
        LeftOffset = LeftOffset + pColumns(i)
        i = i + 1
    Loop
    
    Set UILineitem = New ClsUILineitem
    
    With UILineitem
        .Style = pStyle
        .Left = pLeft + LeftOffset
        .Width = pColumns(Col)
        .Height = pHeight
        .Top = pTop + (Row * pRowOffset)
        .Text = ItemText
        .OnAction = OnAction
    End With
    Me.AddItem UILineitem
End Sub

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the Lineitems sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Lineitem As ClsUILineitem
        
    For Each Lineitem In PUILineitems
        Lineitem.ShpLineitem.ZOrder msoBringToFront
        Lineitem.ReOrder
    Next
    
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIScreen
    Dim tmp As ClsUIScreen

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set PUILineitems = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    
    Set PUILineitems = Nothing

End Sub

'---------------------------------------------------------------
