VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsUIDashObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsUIShpDashObj
' v0,0 - Initial Version
' v0,1 - Added Style property
'---------------------------------------------------------------
' Date - 13 Nov 17
'===============================================================
' Methods
'---------------------------------------------------------------
' ReOrder - re-orders the shapes so that the frames sit on top of the screen
' Kill - Deletes the shape
'===============================================================

Option Explicit
Private pTop As Integer
Private pLeft As Integer
Private pHeight As Integer
Private pWidth As Integer
Private pName As String
Private pText As String
Private pLocked As Boolean
Private pShpDashObj As Shape
Private pEnumObjType As EnumObjType
Private pVisible As Boolean
Private pStyle As TypeStyle
Private pParent As Long


'---------------------------------------------------------------
Public Property Get Top() As Integer
    Top = pShpDashObj.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
    pTop = vNewValue
    pShpDashObj.Top = pTop + Parent.Top
End Property

'---------------------------------------------------------------
Public Property Get Left() As Integer
    Left = pShpDashObj.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
    pLeft = vNewValue
    pShpDashObj.Left = pLeft + Parent.Left
End Property

'---------------------------------------------------------------
Public Property Get Height() As Integer
    Height = pShpDashObj.Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
    pHeight = vNewValue
    pShpDashObj.Height = pHeight
End Property

'---------------------------------------------------------------
Public Property Get Width() As Integer
    Width = pShpDashObj.Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
    pWidth = vNewValue
    pShpDashObj.Width = pWidth
End Property

'---------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal vNewValue As String)
    pName = vNewValue
    pShpDashObj.Name = pName
End Property

'---------------------------------------------------------------
Public Property Get Text() As String
    pText = pShpDashObj.TextFrame.Characters.Text
    Text = pText
End Property

Public Property Let Text(ByVal vNewValue As String)
    pText = vNewValue
    pShpDashObj.TextFrame.Characters.Text = pText
End Property

'---------------------------------------------------------------
Public Property Get Locked() As Boolean
    pLocked = pShpDashObj.ControlFormat.LockedText = Locked
    Locked = pLocked
End Property

Public Property Let Locked(ByVal vNewValue As Boolean)
    pLocked = vNewValue
    pShpDashObj.Locked = True
    pShpDashObj.ControlFormat.LockedText = Locked
End Property

'---------------------------------------------------------------
Public Property Get ShpDashObj() As Shape
    Set ShpDashObj = pShpDashObj
End Property

Public Property Let ShpDashObj(ByVal vNewValue As Shape)
    Set pShpDashObj = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = pShpDashObj.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    pVisible = vNewValue
    If Not pShpDashObj Is Nothing Then pShpDashObj.Visible = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get EnumObjType() As EnumObjType
    EnumObjType = pEnumObjType
End Property

Public Property Let EnumObjType(ByVal vNewValue As EnumObjType)
    pEnumObjType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Style() As TypeStyle
    Style = pStyle
End Property

Public Property Let Style(vNewValue As TypeStyle)
    pStyle = vNewValue
    
    With pShpDashObj.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = vNewValue.Fill1
        .BackColor.RGB = vNewValue.Fill2
    End With
    
    With pShpDashObj.Line
        .ForeColor.RGB = vNewValue.BorderColour
        .Weight = vNewValue.BorderWidth
        If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
    End With
    
    With pShpDashObj.TextFrame
        .Characters.Font.Bold = vNewValue.FontBold
        .Characters.Font.Name = vNewValue.FontStyle
        .Characters.Font.Color = vNewValue.FontColour
        .Characters.Font.Size = vNewValue.FontSize
        If vNewValue.FontXJust <> 0 Then .HorizontalAlignment = vNewValue.FontXJust
    End With
    
    With pShpDashObj.Shadow
        If vNewValue.Shadow = 0 Then
            .Visible = msoFalse
        Else
            .Type = vNewValue.Shadow
            .Visible = msoCTrue
        End If
    End With
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

' ===============================================================
' Method ReOrder
' re-orders the shapes so that the frames sit on top of the screen
'---------------------------------------------------------------
Public Sub ReOrder()
    
    pShpDashObj.ZOrder msoSendToBack
    
End Sub

' ===============================================================
' Method Kill
' Deletes the shape
'---------------------------------------------------------------
Public Sub Kill()
    
    pShpDashObj.Delete
    
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
    Set pShpDashObj = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
End Sub

