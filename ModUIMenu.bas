Attribute VB_Name = "ModUIMenu"
'===============================================================
' Module ModUIMenu
' v0,0 - Initial Version
' v0,1 - changes to performance mode switching
' v0,2 - Refresh front screen orders after new order placed
' v0,3 - Report1 Button and moved ResetScreen procedures in
' v0,4 - Added Exit Button
' v0,5 - Exit button leaves other workbooks open
'---------------------------------------------------------------
' Date - 05 Jul 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModUIMenu"

' ===============================================================
' BuildStylesMenu
' Builds the UI styles for use on the menu and back drop
' ---------------------------------------------------------------
Public Function BuildStylesMenu() As Boolean
    Const StrPROCEDURE As String = "BuildStylesMenu()"

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
    
    BuildStylesMenu = True

Exit Function

ErrorExit:

    BuildStylesMenu = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildMenu
' Builds the menu using shapes
' ---------------------------------------------------------------
Public Function BuildMenu() As Boolean
    
    Const StrPROCEDURE As String = "BuildMenu()"

    On Error GoTo ErrorHandler
        
    Set MainScreen = New ClsUIScreen
    Set MenuBar = New ClsUIFrame

    If Not BuildBackDrop Then Err.Raise HANDLED_ERROR
    If Not BuildMenuBar Then Err.Raise HANDLED_ERROR
    
    BuildMenu = True
       
Exit Function

ErrorExit:

    BuildMenu = False
    
Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BuildBackDrop
' Builds the background image
' ---------------------------------------------------------------
Private Function BuildBackDrop() As Boolean
    Const StrPROCEDURE As String = "BuildBackDrop()"

    On Error GoTo ErrorHandler

    'Main Screen
    With MainScreen
        .Style = SCREEN_STYLE
        .Name = "Main Screen"
        .Top = 10
        .Left = 10
        .Height = SCREEN_HEIGHT
        .Width = SCREEN_WIDTH
    End With
    
    BuildBackDrop = True

Exit Function

ErrorExit:

    BuildBackDrop = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' BuildMenuBar
' Builds menu on menu bar
' ---------------------------------------------------------------
Private Function BuildMenuBar() As Boolean
    Dim MenuItemText() As String
    Dim MenuItemIcon() As String
    Dim i As Integer
    
    Const StrPROCEDURE As String = "BuildMenuBar()"

    On Error GoTo ErrorHandler
    
     MainScreen.Frames.AddItem MenuBar
   
   'Menubar
    With MenuBar
        .Top = MENUBAR_TOP
        .Left = MENUBAR_LEFT
        .Height = MENUBAR_HEIGHT
        .Width = MENUBAR_WIDTH
        .Name = "MenuBar"
        .Style = MENUBAR_STYLE
        .Header.Visible = False
        .EnableHeader = False
    End With

'    'Logo
'    With Logo
'        .ShpDashObj = ShtMain.Shapes("TEMPLATE - Logo").Duplicate
'        .Name = "Logo"
'        .EnumObjType = ObjImage
'        .Visible = True
'        .Top = LOGO_TOP
'        .Left = LOGO_LEFT
'        .Width = LOGO_WIDTH
'        .Height = LOGO_HEIGHT
'    End With

'    MenuBar.DashObs.AddItem Logo
    

    'menu
    With MenuBar.Menu
        .Top = MENU_TOP
        .Left = MENU_LEFT
    End With

    'Menu Items
    MenuItemText() = Split(MENUITEM_TEXT, ":")
    MenuItemIcon() = Split(MENUITEM_ICONS, ":")

    For i = 0 To MENUITEM_COUNT - 1

        Set MenuItem = New ClsUIMenuItem
    
        With MenuItem
            .SelectStyle = MENUITEM_SET_STYLE
            .UnSelectStyle = MENUITEM_UNSET_STYLE
            .Height = MENUITEM_HEIGHT
            .Width = MENUITEM_WIDTH
            .Text = MenuItemText(i)
            .Name = "MenuItem - " & .Text
            .OnAction = "'ModUIMenu.ProcessBtnPress(" & i + 1 & ")'"
            .Icon = ShtMain.Shapes(MenuItemIcon(i)).Duplicate

            MenuBar.Menu.AddItem MenuItem

            .Top = MENU_TOP + (i * .Height) - i
            .Left = .Left
            .Selected = False

            With .Icon
                .Visible = True
                .Name = "Icon - " & MenuItem.Text
                .Left = MenuItem.Left + MENUITEM_ICON_LEFT
                .Top = MenuItem.Top + MENUITEM_ICON_TOP
            End With
        End With
    Next
    
    Set MenuItem = Nothing

    BuildMenuBar = True

Exit Function

ErrorExit:

    Set MenuItem = Nothing
    
    BuildMenuBar = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ProcessBtnPress
' Receives all button presses and processes
' ---------------------------------------------------------------
Public Function ProcessBtnPress(ButtonNo As EnumBtnNo) As Boolean
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "ProcessBtnPress()"

    On Error GoTo ErrorHandler
    
        If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
        
Restart:
        ShtMain.Unprotect
        
        Application.StatusBar = ""
        
        If ButtonNo < 6 And ButtonNo = MENU_ITEM_SEL Then Exit Function
        
        MENU_ITEM_SEL = ButtonNo
        
        Select Case ButtonNo
        
            Case EnumMyStation
                
                [menuitemno] = 1

                If Not ResetScreen Then Err.Raise HANDLED_ERROR
                If Not ModUIMainScreen.BuildMainScreen Then Err.Raise HANDLED_ERROR
                                
                With MenuBar
                    .Menu(1).Selected = True
                    .Menu(2).Selected = False
                    .Menu(3).Selected = False
                    .Menu(4).Selected = False
                    .Menu(5).Selected = False
                    .Menu(6).Selected = False
                End With
                
            Case EnumStores
                
                [menuitemno] = 2
                
'                If CurrentUser.AccessLvl < StoresLvl_2 Then Err.Raise ACCESS_DENIED
'
'                If Not ResetScreen Then Err.Raise HANDLED_ERROR
'                If Not ModUIStoresScreen.BuildStoresScreen Then Err.Raise HANDLED_ERROR
'
'                ShtMain.Unprotect
'
'                With MenuBar
'                    .Menu(1).Selected = False
'                    .Menu(2).Selected = True
'                    .Menu(3).Selected = False
'                    .Menu(4).Selected = False
'                    .Menu(5).Selected = False
'                    .Menu(6).Selected = False
'                End With
'
'            Case EnumReports
'
'                ShtMain.Unprotect
'
'                [menuitemno] = 3
'
'                If CurrentUser.AccessLvl < ManagerLvl_4 Then Err.Raise ACCESS_DENIED
'
'                ModLibrary.PerfSettingsOn
'
'                If Not ResetScreen Then Err.Raise HANDLED_ERROR
'                If Not ModUIReporting.BuildReporting Then Err.Raise HANDLED_ERROR
'
'                ShtMain.ClearOrderList
'
'                With MenuBar
'                    .Menu(1).Selected = False
'                    .Menu(2).Selected = False
'                    .Menu(3).Selected = True
'                    .Menu(4).Selected = False
'                    .Menu(5).Selected = False
'                    .Menu(6).Selected = False
'                End With
'
'                ModLibrary.PerfSettingsOff
'
'            Case EnumMyProfile
'
'                ShtMain.Unprotect
'
'                [menuitemno] = 4
'
'                ModLibrary.PerfSettingsOn
'
'                If Not ResetScreen Then Err.Raise HANDLED_ERROR
'
'                ShtMain.ClearOrderList
'
'                With MenuBar
'                    .Menu(1).Selected = False
'                    .Menu(2).Selected = False
'                    .Menu(3).Selected = False
'                    .Menu(4).Selected = True
'                    .Menu(5).Selected = False
'                    .Menu(6).Selected = False
'                End With
'
'                ModLibrary.PerfSettingsOff
'
'            Case EnumSupport
'
'                ShtMain.Unprotect
'
'                [menuitemno] = 5
'
'                ModLibrary.PerfSettingsOn
'
'                If Not ResetScreen Then Err.Raise HANDLED_ERROR
'                If Not ModUISupportScreen.BuildSupportScreen Then Err.Raise HANDLED_ERROR
'
'                With MenuBar
'                    .Menu(1).Selected = False
'                    .Menu(2).Selected = False
'                    .Menu(3).Selected = False
'                    .Menu(4).Selected = False
'                    .Menu(5).Selected = True
'                    .Menu(6).Selected = False
'                End With
'
'                ModLibrary.PerfSettingsOff
'
'            Case EnumNewOrder
'
'                If Not FrmOrder.ShowForm Then Err.Raise HANDLED_ERROR
'                If Not ModUIMainScreen.RefreshMyOrderList Then Err.Raise HANDLED_ERROR
'                If Not ModUIMainScreen.RefreshRecentOrderList Then Err.Raise HANDLED_ERROR
'
'            Case EnumExit
'
'                ModLibrary.PerfSettingsOn
'
'                Response = MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME)
'
'                If Response = 6 Then
'
'                    If Workbooks.Count = 1 Then
'                        With Application
'                            .DisplayAlerts = True
'                            .Quit
'                            .DisplayAlerts = False
'                        End With
'                    Else
'                        ActiveWorkbook.Close savechanges:=False
'                    End If
'
'                End If
'
'                ModLibrary.PerfSettingsOff
'
        End Select
        
        ShtMain.Protect
        
GracefulExit:
    
    ModLibrary.PerfSettingsOff

    ProcessBtnPress = True

Exit Function

ErrorExit:

    Application.DisplayAlerts = True

    ProcessBtnPress = False

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        CustomErrorHandler Err.Number
         If Err.Number = SYSTEM_RESTART Then
            Resume Restart
        Else
            Resume GracefulExit
        End If
    End If
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ResetScreen
' Functions for graceful close down of system
' ---------------------------------------------------------------
Public Function ResetScreen() As Boolean
    Dim Frame As ClsUIFrame
    Dim UILineitem As ClsUILineitem
    Dim DashObj As ClsUIDashObj
    Dim MenuItem As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "ResetScreen()"

    On Error Resume Next
        
    ShtMain.Unprotect
    
    For Each Frame In MainScreen.Frames
        If Frame.Name <> "MenuBar" Then
            For Each DashObj In Frame.DashObs
'                Debug.Print DashObj.Name
                Frame.DashObs.RemoveItem DashObj.Name
                DashObj.ShpDashObj.Delete
                Set DashObj = Nothing
            Next
            
            For Each UILineitem In Frame.Lineitems
'                Debug.Print UILineitem.Name
                Frame.Lineitems.RemoveItem UILineitem.Name
                UILineitem.ShpLineitem.Delete
                Set UILineitem = Nothing
            Next
            
            For Each MenuItem In Frame.Menu
                Frame.Menu.RemoveItem MenuItem.Name
                MenuItem.ShpMenuItem.Delete
                MenuItem.Icon.Delete
                Set MenuItem = Nothing
            Next
            
            Frame.Header.Icon.Delete
            Frame.Header.ShpHeader.Delete
            Set Frame.Header = Nothing
            
            MainScreen.Frames.RemoveItem Frame.Name
            Frame.ShpFrame.Delete
            Set Frame = Nothing
            
        End If
    Next
    
    For Each MenuItem In MainScreen.Menu
        MainScreen.Menu.RemoveItem MenuItem.Name
        MenuItem.ShpMenuItem.Delete
        MenuItem.Icon.Delete
        Set MenuItem = Nothing
    Next
        
    ResetScreen = True
        
Exit Function

ErrorExit:

    ResetScreen = False

Exit Function

ErrorHandler:
    
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



