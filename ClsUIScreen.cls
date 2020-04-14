VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIScreen
' v0,0 - Initial Version
'---------------------------------------------------------------
' Date - 06 Feb 17
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
Private pStyle As TypeStyle
Private pMenu As ClsUIMenu
Private pShpScreen As Shape
Private pFrames As ClsUIFrames
Private pVisible As Boolean

'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpScreen.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpScreen.Top = pTop
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpScreen.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpScreen.Left = pLeft
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpScreen.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpScreen.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpScreen.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpScreen.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pShpScreen.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpScreen.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get ShpScreen() As Shape
    Set ShpScreen = pShpScreen
End Property

Public Property Let ShpScreen(ByVal vNewValue As Shape)
    Set pShpScreen = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Frames() As ClsUIFrames
    Set Frames = pFrames
End Property

Public Property Let Frames(ByVal vNewValue As ClsUIFrames)
    Set pFrames = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Menu() As ClsUIMenu
    Set Menu = pMenu
End Property

Public Property Let Menu(ByVal vNewValue As ClsUIMenu)
    Set pMenu = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
    
    With pShpScreen.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = pStyle.Fill1
        .BackColor.RGB = pStyle.Fill2
    End With
    
    With pShpScreen.Line
        .ForeColor.RGB = pStyle.BorderColour
        .Weight = pStyle.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpScreen.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = pStyle.FontStyle
        .Characters.Font.Color = pStyle.FontColour
        .Characters.Font.Size = pStyle.FontSize
        If pStyle.FontXJust <> 0 Then .HorizontalAlignment = pStyle.FontXJust
    End With
    
    With pShpScreen.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpScreen.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    pShpScreen.Visible = pVisible
End Property

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    Dim Frame As ClsUIFrame
    
    pShpScreen.ZOrder msoSendToBack
    
    For Each Frame In pFrames
        Frame.ReOrder
    Next
    
    BtnNewOrder.ReOrder
    
End Sub

' ===============================================================
' Method Kill
' Deletes the shape
'---------------------------------------------------------------
Public Sub Kill()
    Dim Frame As ClsUIFrame
    
    For Each Frame In pFrames
        'debug.print Frame.Name
        If Frame.Name <> "MenuBar" Then Frame.ShpFrame.Delete
        Set Frame = Nothing
    Next
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pFrames = New ClsUIFrames
    Set pMenu = New ClsUIMenu
    
    pFrames.SetParent ObjPtr(Me)
    pMenu.SetParent ObjPtr(Me)
    
    Set pShpScreen = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    pFrames.SetParent 0
    pMenu.SetParent 0
    
    Set pFrames = Nothing
    Set pMenu = Nothing
    
    pShpScreen.Delete
End Sub

'---------------------------------------------------------------


