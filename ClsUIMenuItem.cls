VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIMenuItem
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Feb 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so that the frames sit on top of the screen
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pText As String
Private pIcon As Shape
Private pSelectStyle As TypeStyle
Private pUnSelectStyle As TypeStyle
Private pShpMenuItem As Shape
Private pSelected As Boolean
Private pVisible As Boolean
Private pOnAction As String
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpMenuItem.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    Dim ParentOffset As Integer
    
    pTop = vNewValue
    
    If Not Parent Is Nothing Then ParentOffset = Parent.Top
    
    pShpMenuItem.Top = pTop + ParentOffset
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpMenuItem.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    Dim ParentOffset As Integer
    
    pLeft = vNewValue
    
    If Not Parent Is Nothing Then ParentOffset = Parent.Left
    
    pShpMenuItem.Left = pLeft + ParentOffset
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpMenuItem.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpMenuItem.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpMenuItem.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpMenuItem.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpMenuItem.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpMenuItem.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    Text = pShpMenuItem.TextFrame.Characters.Caption
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpMenuItem.TextFrame.Characters.Caption = pText
End Property

'---------------------------------------------------------------
Public Property Get Icon() As Shape
    Set Icon = pIcon
End Property

Public Property Let Icon(ByVal vNewValue As Shape)
    Set pIcon = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get SelectStyle() As TypeStyle
    SelectStyle = pSelectStyle
End Property

Public Property Let SelectStyle(vNewValue As TypeStyle)
    pSelectStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get UnSelectStyle() As TypeStyle
    UnSelectStyle = pUnSelectStyle
End Property

Public Property Let UnSelectStyle(vNewValue As TypeStyle)
    pUnSelectStyle = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ShpMenuItem() As Shape
    Set ShpMenuItem = pShpMenuItem
End Property

Public Property Let ShpMenuItem(ByVal vNewValue As Shape)
    Set pShpMenuItem = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Selected() As Boolean
    Selected = pSelected
End Property

Public Property Let Selected(ByVal vNewValue As Boolean)
    Dim Style As TypeStyle
    
    pSelected = vNewValue
        
    If pSelected Then
        Style = pSelectStyle
    Else
        Style = pUnSelectStyle
    End If
    
    With pShpMenuItem.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = Style.Fill1
        .BackColor.RGB = Style.Fill2
    End With
    
    With pShpMenuItem.Line
        .ForeColor.RGB = Style.BorderColour
        .Weight = Style.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpMenuItem.TextFrame
        .Characters.Font.Bold = Style.FontBold
        .Characters.Font.Name = Style.FontStyle
        .Characters.Font.Color = Style.FontColour
        .Characters.Font.Size = Style.FontSize
        If Style.FontXJust <> 0 Then .HorizontalAlignment = Style.FontXJust
        If Style.FontYJust <> 0 Then .VerticalAlignment = Style.FontYJust
    End With
    
    With pShpMenuItem.Shadow
        If Style.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = Style.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pVisible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    pShpMenuItem.Visible = vNewValue
    If Not pIcon Is Nothing Then pIcon.Visible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get OnAction() As String
    OnAction = pShpMenuItem.OnAction
End Property

Public Property Let OnAction(ByVal vNewValue As String)
    pShpMenuItem.OnAction = vNewValue
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    pShpMenuItem.ZOrder msoBringToFront
    If Not pIcon Is Nothing Then pIcon.ZOrder msoBringToFront
End Sub

'---------------------------------------------------------------
Public Property Get Parent() As ClsUIMenu
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsUIMenu
    Dim tmp As ClsUIMenu

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pShpMenuItem = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
End Sub

'---------------------------------------------------------------
