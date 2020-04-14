VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUILineitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUiLineitem
' v0,0 - Initial Version
' v0,1 - Remove dashobjs
'---------------------------------------------------------------
' Date - 06 Apr 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so the objects sit in the right order
' Text - Adds text to the line item
'===============================================================

Option Explicit
Private pTop As Integer
Private pHeight As Integer
Private pLeft As Integer
Private pWidth As Integer
Private pText As String
Private pStyle As TypeStyle
Private pName As String
Private pMenu As ClsUIMenu
Private pOnAction As String
Private pShpLineitem As Shape
Private pVisible As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpLineitem.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpLineitem.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpLineitem.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpLineitem.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpLineitem.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpLineitem.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpLineitem.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpLineitem.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pShpLineitem.TextFrame.Characters.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpLineitem.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpLineitem.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpLineitem.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal vNewValue As ClsUIMenu)
    Set pMenu = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpLineitem.OnAction
End Property

Public Property Let OnAction(ByVal vNewValue As String)
    pOnAction = vNewValue
    pShpLineitem.OnAction = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpLineitem.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    Dim DashObj As ClsUIDashObj
    
    pVisible = vNewValue
    
    pShpLineitem.Visible = pVisible
       
End Property

'---------------------------------------------------------------
Public Property Get ShpLineitem() As Shape
    Set ShpLineitem = pShpLineitem
End Property

Public Property Let ShpLineitem(ByVal vNewValue As Shape)
    Set pShpLineitem = vNewValue
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
    
    With pShpLineitem.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With pShpLineitem.Line
        .ForeColor.RGB = vNewValue.BorderColour
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpLineitem.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
    End With
    
    With pShpLineitem.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsUILineitems
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function


' ===============================================================
' Method ReOrder
' re-orders the shapes so that the Lineitems sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim DashObj As ClsUIDashObj
    
    pShpLineitem.ZOrder msoBringToFront
    
    pMenu.ReOrder
End Sub

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUILineitem
    Dim tmp As ClsUILineitem

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()

    Set pShpLineitem = ShtMain.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    

End Sub

'---------------------------------------------------------------
