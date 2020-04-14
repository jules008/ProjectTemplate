Attribute VB_Name = "ModUIMainScreen"
'===============================================================
' Module ModUIMainScreen
' v0,0 - Initial Version
' v0,1 - added performance mode switching
' v0,22 - Build Right frame order list
' v0,3 - Build Left Frame
' v0,4 - Turned performance settings off on error
' v0,5 - Refresh both order panels of order delete
' v0,6 - Moved ResetScreen to main Menu procedure
' v0,7 - Removed hard numbering from buttons
' v0,8 - Added Release Notes Frame
'---------------------------------------------------------------
' Date - 13 Nov 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIMainScreen"

' ===============================================================
' BuildStylesMainScreen
' Builds the UI styles for use on the screen
' ---------------------------------------------------------------
Public Function BuildStylesMainScreen() As Boolean
    Const StrPROCEDURE As String = "BuildStylesMainScreen()"

    On Error GoTo ErrorHandler

    With SCREEN_STYLE
        .BorderWidth = SCREEN_BORDER_WIDTH
        .Fill1 = SCREEN_FILL_1
        .Fill2 = SCREEN_FILL_2
        .Shadow = SCREEN_SHADOW
    End With
    
    With MENUBAR_STYLE
        .BorderWidth = MENUBAR_BORDER_WIDTH
        .Fill1 = MENUBAR_FILL_1
        .Fill2 = MENUBAR_FILL_2
        .Shadow = MENUBAR_SHADOW
    End With
    
    With MENUITEM_UNSET_STYLE
        .BorderWidth = MENUITEM_UNSET_BORDER_WIDTH
        .Fill1 = MENUITEM_UNSET_FILL_1
        .Fill2 = MENUITEM_UNSET_FILL_2
        .Shadow = MENUITEM_UNSET_SHADOW
        .FontStyle = MENUITEM_UNSET_FONT_STYLE
        .FontSize = MENUITEM_UNSET_FONT_SIZE
        .FontColour = MENUITEM_UNSET_FONT_COLOUR
        .FontXJust = MENUITEM_UNSET_FONT_X_JUST
        .FontYJust = MENUITEM_UNSET_FONT_Y_JUST
    End With

    With MENUITEM_SET_STYLE
        .BorderWidth = MENUITEM_SET_BORDER_WIDTH
        .Fill1 = MENUITEM_SET_FILL_1
        .Fill2 = MENUITEM_SET_FILL_2
        .Shadow = MENUITEM_SET_SHADOW
        .FontStyle = MENUITEM_SET_FONT_STYLE
        .FontSize = MENUITEM_SET_FONT_SIZE
        .FontColour = MENUITEM_SET_FONT_COLOUR
        .FontXJust = MENUITEM_SET_FONT_X_JUST
        .FontYJust = MENUITEM_SET_FONT_Y_JUST
    End With
    
    With MAIN_FRAME_STYLE
        .BorderWidth = MAIN_FRAME_BORDER_WIDTH
        .Fill1 = MAIN_FRAME_FILL_1
        .Fill2 = MAIN_FRAME_FILL_2
        .Shadow = MAIN_FRAME_SHADOW
    End With
    
    With HEADER_STYLE
        .BorderWidth = HEADER_BORDER_WIDTH
        .Fill1 = HEADER_FILL_1
        .Fill2 = HEADER_FILL_2
        .Shadow = HEADER_SHADOW
        .FontStyle = HEADER_FONT_STYLE
        .FontSize = HEADER_FONT_SIZE
        .FontBold = HEADER_FONT_BOLD
        .FontColour = HEADER_FONT_COLOUR
        .FontXJust = HEADER_FONT_X_JUST
        .FontYJust = HEADER_FONT_Y_JUST
    End With
    
    With BTN_NEWORDER_STYLE
        .BorderWidth = BTN_NEWORDER_BORDER_WIDTH
        .Fill1 = BTN_NEWORDER_FILL_1
        .Fill2 = BTN_NEWORDER_FILL_2
        .Shadow = BTN_NEWORDER_SHADOW
        .FontStyle = BTN_NEWORDER_FONT_STYLE
        .FontSize = BTN_NEWORDER_FONT_SIZE
        .FontBold = BTN_NEWORDER_FONT_BOLD
        .FontColour = BTN_NEWORDER_FONT_COLOUR
        .FontXJust = BTN_NEWORDER_FONT_X_JUST
        .FontYJust = BTN_NEWORDER_FONT_Y_JUST
    End With
    
    With GENERIC_BUTTON
        .BorderWidth = GENERIC_BUTTON_BORDER_WIDTH
        .Fill1 = GENERIC_BUTTON_FILL_1
        .Fill2 = GENERIC_BUTTON_FILL_2
        .Shadow = GENERIC_BUTTON_SHADOW
        .FontStyle = GENERIC_BUTTON_FONT_STYLE
        .FontSize = GENERIC_BUTTON_FONT_SIZE
        .FontBold = GENERIC_BUTTON_FONT_BOLD
        .FontColour = GENERIC_BUTTON_FONT_COLOUR
        .FontXJust = GENERIC_BUTTON_FONT_X_JUST
        .FontYJust = GENERIC_BUTTON_FONT_Y_JUST
    End With
    
    With GENERIC_LINEITEM
        .BorderWidth = GENERIC_LINEITEM_BORDER_WIDTH
        .Fill1 = GENERIC_LINEITEM_FILL_1
        .Fill2 = GENERIC_LINEITEM_FILL_2
        .Shadow = GENERIC_LINEITEM_SHADOW
        .FontStyle = GENERIC_LINEITEM_FONT_STYLE
        .FontSize = GENERIC_LINEITEM_FONT_SIZE
        .FontBold = GENERIC_LINEITEM_FONT_BOLD
        .FontColour = GENERIC_LINEITEM_FONT_COLOUR
        .FontXJust = GENERIC_LINEITEM_FONT_X_JUST
        .FontYJust = GENERIC_LINEITEM_FONT_Y_JUST
    End With

    With GENERIC_LINEITEM_HEADER
        .BorderWidth = GENERIC_LINEITEM_HEADER_BORDER_WIDTH
        .Fill1 = GENERIC_LINEITEM_HEADER_FILL_1
        .Fill2 = GENERIC_LINEITEM_HEADER_FILL_2
        .Shadow = GENERIC_LINEITEM_HEADER_SHADOW
        .FontStyle = GENERIC_LINEITEM_HEADER_FONT_STYLE
        .FontSize = GENERIC_LINEITEM_HEADER_FONT_SIZE
        .FontBold = GENERIC_LINEITEM_HEADER_FONT_BOLD
        .FontColour = GENERIC_LINEITEM_HEADER_FONT_COLOUR
        .FontXJust = GENERIC_LINEITEM_HEADER_FONT_X_JUST
        .FontYJust = GENERIC_LINEITEM_HEADER_FONT_Y_JUST
    End With
    
    With TRANSPARENT_TEXT_BOX
        .BorderWidth = TRANSPARENT_TEXT_BOX_BORDER_WIDTH
        .Fill1 = TRANSPARENT_TEXT_BOX_FILL_1
        .Fill2 = TRANSPARENT_TEXT_BOX_FILL_2
        .Shadow = TRANSPARENT_TEXT_BOX_SHADOW
        .FontStyle = TRANSPARENT_TEXT_BOX_FONT_STYLE
        .FontSize = TRANSPARENT_TEXT_BOX_FONT_SIZE
        .FontBold = TRANSPARENT_TEXT_BOX_FONT_BOLD
        .FontColour = TRANSPARENT_TEXT_BOX_FONT_COLOUR
        .FontXJust = TRANSPARENT_TEXT_BOX_FONT_X_JUST
        .FontYJust = TRANSPARENT_TEXT_BOX_FONT_Y_JUST
    End With
    

    BuildStylesMainScreen = True

Exit Function
    
    
ErrorExit:

    BuildStylesMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainFrame
' Builds main frame at top of screen
' ---------------------------------------------------------------
Private Function BuildMainFrame() As Boolean
    Const StrPROCEDURE As String = "BuildMainFrame()"

    On Error GoTo ErrorHandler

    Set MainFrame = New ClsUIFrame
    
    'add main frame
    With MainFrame
        .Name = "Main Frame"
        MainScreen.Frames.AddItem MainFrame
            
        .Top = MAIN_FRAME_TOP
        .Left = MAIN_FRAME_LEFT
        .Width = MAIN_FRAME_WIDTH
        .Height = MAIN_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Main Frame Header"
'            .Text = "Allocations - " & CurrentUser.Station.Name
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Alocations").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
    End With
    
    
    BuildMainFrame = True

Exit Function

ErrorExit:

    BuildMainFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMainScreen
' Builds the display using shapes
' ---------------------------------------------------------------
Public Function BuildMainScreen() As Boolean
    
    Const StrPROCEDURE As String = "BuildMainScreen()"

    On Error GoTo ErrorHandler
    
    ModLibrary.PerfSettingsOn
    
    If Not BuildMainFrame Then Err.Raise HANDLED_ERROR
    If Not BuildLeftFrame Then Err.Raise HANDLED_ERROR
    If Not BuildRightFrame Then Err.Raise HANDLED_ERROR
    If Not BuildNewOrderBtn Then Err.Raise HANDLED_ERROR
    
    MainScreen.ReOrder
    
    ModLibrary.PerfSettingsOff
                    
    BuildMainScreen = True
       
Exit Function

ErrorExit:
    
    ModLibrary.PerfSettingsOff

    BuildMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildLeftFrame
' Builds Left frame at top of screen
' ---------------------------------------------------------------
Private Function BuildLeftFrame() As Boolean
    Const StrPROCEDURE As String = "BuildLeftFrame()"

    On Error GoTo ErrorHandler

    Set LeftFrame = Nothing
    Set LeftFrame = New ClsUIFrame
    
    'add Left frame
    With LeftFrame
        .Name = "Left Frame"
        MainScreen.Frames.AddItem LeftFrame
        
        .Top = LEFT_FRAME_TOP
        .Left = LEFT_FRAME_LEFT
        .Width = LEFT_FRAME_WIDTH
        .Height = LEFT_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Left Frame Header"
            .Text = "Recent Orders"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Left_Frame").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        
        End With
        
        With .Lineitems
            .NoColumns = MY_ORDER_LINEITEM_NOCOLS
            .Top = MY_ORDER_LINEITEM_TOP
            .Left = MY_ORDER_LINEITEM_LEFT
            .Height = MY_ORDER_LINEITEM_HEIGHT
            .Columns = MY_ORDER_LINEITEM_COL_WIDTHS
            .RowOffset = MY_ORDER_LINEITEM_ROWOFFSET
                
        End With
        If Not RefreshRecentOrderList Then Err.Raise HANDLED_ERROR
        .ReOrder
        
    End With

    
    BuildLeftFrame = True

Exit Function

ErrorExit:

    BuildLeftFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildRightFrame
' Builds Right frame at top of screen
' ---------------------------------------------------------------
Private Function BuildRightFrame() As Boolean
    Const StrPROCEDURE As String = "BuildRightFrame()"

    On Error GoTo ErrorHandler

    Set RightFrame = Nothing
    Set RightFrame = New ClsUIFrame
    
    'add Right frame
    With RightFrame
        .Name = "Right Frame"
        MainScreen.Frames.AddItem RightFrame
        
        .Top = RIGHT_FRAME_TOP
        .Left = RIGHT_FRAME_LEFT
        .Width = RIGHT_FRAME_WIDTH
        .Height = RIGHT_FRAME_HEIGHT
        .Style = MAIN_FRAME_STYLE
        .EnableHeader = True

        With .Header
            .Top = .Parent.Top
            .Left = .Parent.Left
            .Width = .Parent.Width
            .Height = HEADER_HEIGHT
            .Name = "Right Frame Header"
            .Text = "My Orders"
            .Style = HEADER_STYLE
            .Icon = ShtMain.Shapes("TEMPLATE - Icon_Right_Frame").Duplicate
            .Icon.Top = .Parent.Top + HEADER_ICON_TOP
            .Icon.Left = .Parent.Left + .Parent.Width - .Icon.Width - HEADER_ICON_RIGHT
            .Icon.Name = .Parent.Name & " Icon"
            .Icon.Visible = msoCTrue
        End With
        
        With .Lineitems
            .NoColumns = MY_ORDER_LINEITEM_NOCOLS
            .Top = MY_ORDER_LINEITEM_TOP
            .Left = MY_ORDER_LINEITEM_LEFT
            .Height = MY_ORDER_LINEITEM_HEIGHT
            .Columns = MY_ORDER_LINEITEM_COL_WIDTHS
            .RowOffset = MY_ORDER_LINEITEM_ROWOFFSET
                
        End With
        If Not RefreshMyOrderList Then Err.Raise HANDLED_ERROR
        .ReOrder
    End With

    
    BuildRightFrame = True

Exit Function

ErrorExit:

    BuildRightFrame = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildNewOrderBtn
' Adds the new order button to the main screen
' ---------------------------------------------------------------
Private Function BuildNewOrderBtn() As Boolean

    Const StrPROCEDURE As String = "BuildNewOrderBtn()"

    On Error GoTo ErrorHandler
    
    Set BtnNewOrder = New ClsUIMenuItem

    With BtnNewOrder
        .Height = BTN_NEWORDER_HEIGHT
        .Left = BTN_NEWORDER_LEFT
        .Top = BTN_NEWORDER_TOP
        .Width = BTN_NEWORDER_WIDTH
        .Name = "New Order Button"
        .OnAction = "'moduimenu.ProcessBtnPress(" & EnumNewOrder & ")'"
        .UnSelectStyle = BTN_NEWORDER_STYLE
        .Selected = False
        .Text = "New Order    "
        .Icon = ShtMain.Shapes("TEMPLATE - Icon_New_Order").Duplicate
        .Icon.Left = .Left + 290
        .Icon.Top = .Top + 16
        .Icon.Name = "New_Order_Button"
        .Icon.Visible = msoCTrue
    End With
    
    MainScreen.Menu.AddItem BtnNewOrder
    
    BuildNewOrderBtn = True

Exit Function

ErrorExit:

    BuildNewOrderBtn = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DestroyMainScreen
' Destroys the main screen objects
' ---------------------------------------------------------------
Public Function DestroyMainScreen() As Boolean
    Dim Frame As ClsUIFrame
    
    Const StrPROCEDURE As String = "DestroyMainScreen()"

    On Error GoTo ErrorHandler
    
    Set Frame = New ClsUIFrame
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            MainScreen.Frames.RemoveItem Frame.Name
        End If
    Next
        
    If Not MainFrame Is Nothing Then MainFrame.Visible = False
    If Not LeftFrame Is Nothing Then LeftFrame.Visible = False
    If Not RightFrame Is Nothing Then RightFrame.Visible = False
    If Not BtnNewOrder Is Nothing Then BtnNewOrder.Visible = False
    
    Set MainFrame = Nothing
    Set LeftFrame = Nothing
    Set RightFrame = Nothing
    Set BtnNewOrder = Nothing
    
    Set Frame = Nothing
    
    DestroyMainScreen = True
       
Exit Function

ErrorExit:

    Set Frame = Nothing
    
    DestroyMainScreen = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RefreshMyOrderList
' Populates right frame with my orders
' ---------------------------------------------------------------
Public Function RefreshMyOrderList() As Boolean
    Dim OrderNo As Integer
    Dim OrderDate As String
    Dim AssignedTo As String
    Dim OrderStatus As String
'    Dim Orders As ClsOrders
    Dim RstOrder As Recordset
    Dim Lineitem As ClsUILineitem
    Dim StrOnAction As String
    Dim i As Integer
    Dim RowTitles() As String
    
    Const StrPROCEDURE As String = "RefreshMyOrderList()"

    On Error GoTo ErrorHandler
    
'    Set Orders = New ClsOrders

    ShtMain.Unprotect
    
    ModLibrary.PerfSettingsOn
    
    With RightFrame
        For Each Lineitem In .Lineitems
            .Lineitems.RemoveItem Lineitem.Name
            Lineitem.ShpLineitem.Delete
            Set Lineitem = Nothing
        Next

        ReDim RowTitles(0 To MY_ORDER_LINEITEM_NOCOLS - 1)
        RowTitles = Split(MY_ORDER_LINEITEM_TITLES, ":")

        .Lineitems.Style = GENERIC_LINEITEM_HEADER
        
        For i = 0 To MY_ORDER_LINEITEM_NOCOLS - 1
            .Lineitems.Text 0, i, RowTitles(i), False
        Next
        
        .Lineitems.Style = GENERIC_LINEITEM

    End With

'    Set RstOrder = Orders.MyOrders

    i = 1
    With RstOrder
        Do While Not .EOF
            With RightFrame.Lineitems
                If Not IsNull(RstOrder!Order_No) Then OrderNo = RstOrder!Order_No Else OrderNo = 0
                If Not IsNull(RstOrder!Order_Date) Then OrderDate = RstOrder!Order_Date Else OrderDate = ""
                If Not IsNull(RstOrder!Assigned_To) Then AssignedTo = RstOrder!Assigned_To Else AssignedTo = ""
                If Not IsNull(RstOrder!Status) Then OrderStatus = RstOrder!Status Else OrderStatus = ""
                
                StrOnAction = "'ModUIMainScreen.OpenOrder(" & OrderNo & ")'"
                
                .Text i, 0, CStr(OrderNo), StrOnAction
                .Text i, 1, Format(OrderDate, "dd mmm yy"), StrOnAction
                .Text i, 2, AssignedTo, StrOnAction
                .Text i, 3, OrderStatus, StrOnAction
            End With
            .MoveNext
            i = i + 1
            If i > MY_ORDER_MAX_LINES Then Exit Do
        Loop
        
    End With
    
    ModLibrary.PerfSettingsOff
                

    RefreshMyOrderList = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    RefreshMyOrderList = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RefreshRecentOrderList
' Populates right frame with my orders
' ---------------------------------------------------------------
Public Function RefreshRecentOrderList() As Boolean
    Dim OrderNo As Integer
    Dim OrderDate As String
    Dim OrderedBy As String
    Dim OrderStatus As String
'    Dim Orders As ClsOrders
    Dim RstOrder As Recordset
    Dim Lineitem As ClsUILineitem
    Dim StrOnAction As String
    Dim i As Integer
    Dim RowTitles() As String
    
    Const StrPROCEDURE As String = "RefreshRecentOrderList()"

    On Error GoTo ErrorHandler
    
'    Set Orders = New ClsOrders

    ShtMain.Unprotect
    
    ModLibrary.PerfSettingsOn
    
    With LeftFrame
        For Each Lineitem In .Lineitems
            .Lineitems.RemoveItem Lineitem.Name
            Lineitem.ShpLineitem.Delete
            Set Lineitem = Nothing
        Next

        ReDim RowTitles(0 To RCT_ORDER_LINEITEM_NOCOLS - 1)
        RowTitles = Split(RCT_ORDER_LINEITEM_TITLES, ":")

        .Lineitems.Style = GENERIC_LINEITEM_HEADER
        
        For i = 0 To RCT_ORDER_LINEITEM_NOCOLS - 1
            .Lineitems.Text 0, i, RowTitles(i), False
        Next
        
        .Lineitems.Style = GENERIC_LINEITEM

    End With

'    Set RstOrder = Orders.RecentOrders

    i = 1
    With RstOrder
        Do While Not .EOF
            With LeftFrame.Lineitems
                If Not IsNull(RstOrder!Order_No) Then OrderNo = RstOrder!Order_No Else OrderNo = 0
                If Not IsNull(RstOrder!Order_Date) Then OrderDate = RstOrder!Order_Date Else OrderDate = ""
                If Not IsNull(RstOrder!Name) Then OrderedBy = RstOrder!Name Else OrderedBy = ""
                If Not IsNull(RstOrder!Status) Then OrderStatus = RstOrder!Status Else OrderStatus = ""
                
                StrOnAction = "'ModUIMainScreen.OpenOrder(" & OrderNo & ")'"
                
                .Text i, 0, CStr(OrderNo), StrOnAction
                .Text i, 1, Format(OrderDate, "dd mmm yy"), StrOnAction
                .Text i, 2, OrderedBy, StrOnAction
                .Text i, 3, OrderStatus, StrOnAction
            End With
            .MoveNext
            i = i + 1
            If i > RCT_ORDER_MAX_LINES Then Exit Do
        Loop
        
    End With
    
    ModLibrary.PerfSettingsOff
                

    RefreshRecentOrderList = True

Exit Function

ErrorExit:

'    ***CleanUpCode***
    RefreshRecentOrderList = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' OpenOrder
' Opens the selected order form
' ---------------------------------------------------------------
Private Sub OpenOrder(OrderNo As Integer)
    Const StrPROCEDURE As String = "OpenOrder()"
    
'    Dim Order As ClsOrder
    
    On Error GoTo ErrorHandler

'    Set Order = New ClsOrder
    
'    Order.DBGet OrderNo
    
'    If Not FrmDBOrder.ShowForm(Order) Then Err.Raise HANDLED_ERROR
'
'    ModLibrary.PerfSettingsOn
'    ShtMain.Unprotect
'
'    If Not RefreshMyOrderList Then Err.Raise HANDLED_ERROR
'    If Not RefreshRecentOrderList Then Err.Raise HANDLED_ERROR
'
'    ModLibrary.PerfSettingsOff
'    ShtMain.Protect
'
'    Set Order = Nothing

Exit Sub

ErrorExit:

    ModLibrary.PerfSettingsOff
'    Set Order = Nothing
    Terminate
Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

