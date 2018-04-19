Attribute VB_Name = "ModLibrary"
'===============================================================
' Module ModLibrary
' v0,0 - Initial Version
' v0,1 - Added ColourConvert
'---------------------------------------------------------------
' Date - 08 Feb 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModLibrary"

' ===============================================================
' ConvertHoursIntoDecimal
' Converts standard date format into decimal format
' ---------------------------------------------------------------
Public Function ConvertHoursIntoDecimal(TimeIn As Date)
    On Error Resume Next
    
    Dim TB, Result As Single
    
    TB = Split(TimeIn, ":")
    ConvertHoursIntoDecimal = TB(0) + ((TB(1) * 100) / 60) / 100
    
End Function

' ===============================================================
' EndOfMonth
' Returns the number of days in the given month
' ---------------------------------------------------------------
Function EndOfMonth(InputDate As Date) As Variant
    On Error Resume Next
    
    EndOfMonth = Day(DateSerial(Year(InputDate), Month(InputDate) + 1, 0))
End Function

' ===============================================================
' PerfSettingsOn
' turns off system functions to increase performance
' ---------------------------------------------------------------
Public Sub PerfSettingsOn()
    On Error Resume Next
    
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

' ===============================================================
' PerfSettingsOff
' turns system functions back to normal
' ---------------------------------------------------------------
Public Sub PerfSettingsOff()
    On Error Resume Next
        
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ===============================================================
' SpellCheck
' checks spelling on forms
' ---------------------------------------------------------------
Public Sub SpellCheck(ByRef Cntrls As Collection)
    On Error Resume Next
    
    Dim RngSpell As Range
    Dim Cntrl As Control
    
    Set RngSpell = Worksheets(1).Range("A1")
    
    For Each Cntrl In Cntrls
        
        If Left(Cntrl.Name, 3) = "Txt" Then
            Debug.Print Cntrl.Name
            RngSpell = Cntrl
            RngSpell.CheckSpelling
            Cntrl = RngSpell
        End If
    Next
    
End Sub

' ===============================================================
' RecordsetPrint
' sends contents of recordset to debug window
' ---------------------------------------------------------------
Public Sub RecordsetPrint(RST As Recordset)
    On Error Resume Next
    
    Dim DBString As String
    Dim RSTField As Field
    Dim i As Integer

    ReDim AyFields(RST.Fields.Count)
    
    Do Until RST.EOF
        For i = 0 To RST.Fields.Count - 1
             DBString = DBString & RST.Fields(i).Value & ", "
        Next
        RST.MoveNext
        Debug.Print DBString
        DBString = ""
    Loop

End Sub

' ===============================================================
' PrintPDF
' Prints sent worksheet as a PDF
' ---------------------------------------------------------------
Public Sub PrintPDF(WSheet As Worksheet, PathAndFileName As String)
    On Error Resume Next
    
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler
    
    strFile = PathAndFileName & ".pdf"
    
    WSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler

End Sub

' ===============================================================
' CopyTextToClipboard
' Sends string to clipboard for pasting
' ---------------------------------------------------------------
Sub CopyTextToClipboard()

    Dim obj As New DataObject
    Dim Txt As String
    
    Txt = Chr(9) & "This was copied to the clipboard using VBA!" & Chr(13) & "New Line"
    obj.SetText Txt
    obj.PutInClipboard
    
    MsgBox "There is now text copied to your clipboard!", vbInformation

End Sub

' ===============================================================
' ColourConvert
' Converts RGB colour to long
' ---------------------------------------------------------------

Public Sub ColourConvert()
     Dim Colour1 As Long
     Colour1 = RGB(237, 12, 63)
     
     Debug.Print Colour1

End Sub
