VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIMenuItemu
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - Finds Item from index
' AddItem - Add item to Collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
'===============================================================

Option Explicit
Private pUIMenu As Collection
Private pTop As Integer
Private pLeft As Integer
Private pVisible As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pTop
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pLeft
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    Dim MenuItem As ClsUIMenuItem
    
    pVisible = vNewValue
    
    For Each MenuItem In pUIMenu
        MenuItem.Visible = vNewValue
    Next
    
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
    Set NewEnum = pUIMenu.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUIMenuItem
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUIMenu.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UIMenuItem As ClsUIMenuItem)
    UIMenuItem.SetParent ObjPtr(Me.Parent)
    pUIMenu.Add UIMenuItem, Key:=CStr(UIMenuItem.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUIMenu.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIMenu.Count
End Function

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim MenuItem As ClsUIMenuItem
        
    For Each MenuItem In pUIMenu
        MenuItem.ShpMenuItem.ZOrder msoBringToFront
        MenuItem.ReOrder
    Next
    
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIFrame
    Dim tmp As ClsUIFrame

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pUIMenu = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pUIMenu = Nothing

End Sub

'---------------------------------------------------------------
