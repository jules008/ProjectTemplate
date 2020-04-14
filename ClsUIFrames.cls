VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIFrames
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
' ReOrder - re-orders the shapes so that the frames sit on top of the screen
'===============================================================

Option Explicit
Private pUIFrames As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIScreen
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pUIFrames.[_NewEnum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(Name As Variant) As ClsUIFrame
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pUIFrames.Item(Name)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(UIFrame As ClsUIFrame)
    UIFrame.SetParent ObjPtr(Me.Parent)
    pUIFrames.Add UIFrame, Key:=CStr(UIFrame.Name)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------

Public Sub RemoveItem(Name As Variant)
    pUIFrames.Remove Name
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------

Public Function Count() As Integer
    Count = pUIFrames.Count
End Function

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Frame As ClsUIFrame
        
    For Each Frame In pUIFrames
        Frame.ShpFrame.ZOrder msoBringToFront
        Frame.ReOrder
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
    Set pUIFrames = New Collection

End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    
    Set pUIFrames = Nothing

End Sub

'---------------------------------------------------------------
