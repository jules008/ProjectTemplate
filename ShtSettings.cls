VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub BtmGenClass_Click()
    FrmClassGen.Show

End Sub

Private Sub BtnDBConnect_Click()
    ModDatabase.SelectDB
End Sub

Private Sub BtnGenCollection_Click()
    FrmCollectGen.Show
End Sub

Private Sub BtnGenerateMod_Click()
    FrmModGen.Show
End Sub

Private Sub BtnGenerateProc_Click()
    FrmProcGen.Show
End Sub
