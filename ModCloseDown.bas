Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
' v0,0 - Initial Version
' v0,1 - Delete menu item no on close down
' v0,2 - Made Terminate a function and added Log Off
'---------------------------------------------------------------
' Date - 15 Apr 20
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModCloseDown"

' ===============================================================
' Terminate
' Closedown processing
' ---------------------------------------------------------------
Public Function Terminate() As Boolean
    Dim Frame As ClsUIFrame
    Dim DashObj As ClsUIDashObj
    Dim MenuItem As ClsUIMenuItem
    
    Const StrPROCEDURE As String = "Terminate()"

    On Error Resume Next
        
    ShtMain.Unprotect
    
    CurrentUser.LogUserOff
    
    For Each Frame In MainScreen.Frames
        'debug.print Frame.Name
        For Each DashObj In Frame.DashObs
            'debug.print DashObj.Name
            DashObj.ShpDashObj.Delete
            Set DashObj = Nothing
        Next
        
        For Each MenuItem In Frame.Menu
            'debug.print MenuItem.Name
            MenuItem.ShpMenuItem.Delete
            MenuItem.Icon.Delete
            MenuItem.Badge.Delete
            Set MenuItem = Nothing
        Next
        
        [menuitemno] = ""
        
        Frame.Header.Icon.Delete
        Frame.Header.ShpHeader.Delete
        Set Frame.Header = Nothing
        
        Frame.ShpFrame.Delete
        Set Frame = Nothing
        
    Next
    
    Application.DisplayFullScreen = False
    
    Set MainScreen = Nothing
    
    If Not CurrentUser Is Nothing Then Set CurrentUser = Nothing
    
    ModDatabase.DBTerminate
    DeleteAllShapes

    Terminate = True

Exit Function

ErrorExit:

    If Not CurrentUser Is Nothing Then Set CurrentUser = Nothing

    ModDatabase.DBTerminate
    DeleteAllShapes
    Application.DisplayFullScreen = False
    
    Terminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DeleteAllShapes
' Deletes all shapes on screen except templates
' ---------------------------------------------------------------
Private Sub DeleteAllShapes()
    Dim i As Integer
    
    Const StrPROCEDURE As String = "DeleteAllShapes()"

    On Error Resume Next

    Dim Shp As Shape
    
    For i = ShtMain.Shapes.Count To 1 Step -1
    
        Set Shp = ShtMain.Shapes(i)
        'debug.print i & "/" & ShtMain.Shapes.Count & " " & Shp.Name
        
        If Left(Shp.Name, 8) <> "TEMPLATE" Then Shp.Delete
    Next

End Sub
