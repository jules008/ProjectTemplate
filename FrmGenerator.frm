VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGenerator 
   Caption         =   "UserForm1"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5100
   OleObjectBlob   =   "FrmGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnGenClass_Click()
    FrmClassGen.Show
End Sub

Private Sub BtnGenCollection_Click()
    FrmCollectGen.Show
End Sub

Private Sub BtnGenDBClass_Click()
    FrmDBClassGen.Show
End Sub

Private Sub BtnGenModule_Click()
    FrmModGen.Show
End Sub

Private Sub BtnGenProc_Click()
    FrmProcGen.Show
End Sub
