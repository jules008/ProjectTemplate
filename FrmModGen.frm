VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmModGen 
   Caption         =   "UserForm1"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   OleObjectBlob   =   "FrmModGen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmModGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit


Private Sub BtnGen_Click()
    Dim ModFile
    Dim FileName As String
    Dim i As Integer
    
    ModFile = FreeFile()
    
    FileName = "Mod" & TxtModName & " v0,0.bas"
    
    Open IMPORT_FILE_PATH & FileName For Output As #ModFile
    
    'Write header information
    Print #ModFile, "Attribute VB_Name = """ & "Mod" & TxtModName & """"
    Print #ModFile, "'==============================================================="
    Print #ModFile, "' Module " & "Mod" & TxtModName
    Print #ModFile, "' v0,0 - Initial Version"
    Print #ModFile, "'---------------------------------------------------------------"
    Print #ModFile, "' Date - " & Format(Now, "dd mmm yy")
    Print #ModFile, "'==============================================================="
    Print #ModFile,
    Print #ModFile, "Option Explicit"
    Print #ModFile,
    Print #ModFile, "Private Const StrMODULE As String = """ & "Mod" & TxtModName & """"
    Print #ModFile,
    Close #ModFile
    
    ModProjectInOut.ImportModule FileName
    Hide
End Sub

