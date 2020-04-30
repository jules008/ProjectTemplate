VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()

    On Error GoTo ErrorHandler

    If MainScreen Is Nothing Then Err.Raise SYSTEM_RESTART
        
Restart:
Exit Sub

ErrorHandler:
    
    If Err.Number = SYSTEM_RESTART Then
        CustomErrorHandler Err.Number
       Resume Restart
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
