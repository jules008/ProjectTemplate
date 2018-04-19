Attribute VB_Name = "ModProjectInOut"
'===============================================================
' Module ModProjectInOut
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================
Const EXPORT_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Project Template\Library\Dev\"
Const PROJECT_FILE_NAME As String = "ProjectTemplate"

Public Sub ExportModules()
    Dim ExportYN As Boolean
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim ExportFileName As String
    Dim VBModule As VBIDE.VBComponent
   
    ''' NOTE: This workbook must be open in Excel.
    SourceBookName = ActiveWorkbook.Name
    Set SourceBook = Application.Workbooks(SourceBookName)
    
    If Not Dir(EXPORT_FILE_PATH & "*.*") = "" Then
        Kill EXPORT_FILE_PATH & "*.*"
    End If
    
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        ExportYN = True
        ExportFileName = VBModule.Name

        ''' Concatenate the correct filename for export.
        Select Case VBModule.Type
            Case vbext_ct_ClassModule
                ExportFileName = ExportFileName & ".cls"
            Case vbext_ct_MSForm
                ExportFileName = ExportFileName & ".frm"
            Case vbext_ct_StdModule
                ExportFileName = ExportFileName & ".bas"
            Case vbext_ct_Document
                ExportFileName = ExportFileName & ".cls"
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
'                ExportYN = False
        End Select
        
        If ExportYN Then
            ''' Export the component to a text file.
            VBModule.Export EXPORT_FILE_PATH & ExportFileName
            
        End If
   
    Next VBModule
    
    ThisWorkbook.SaveAs EXPORT_FILE_PATH & "\" & PROJECT_FILE_NAME, 51
    
    Set DlgOpen = Nothing

    MsgBox "Export is ready", vbInformation, APP_NAME
End Sub

Public Sub ImportModules()
    Dim TargetBook As Excel.Workbook
    Dim FSO As Scripting.FileSystemObject
    Dim FileObj As Scripting.File
    Dim TargetBookName As String
    Dim ImportFilePath As String
    Dim ImportFileName As String
    Dim VBModules As VBIDE.VBComponents

    ImportFilePath = ThisWorkbook.Path
    
    ''' NOTE: This workbook must be open in Excel.
    TargetBookName = ActiveWorkbook.Name
    Set TargetBook = Application.Workbooks(TargetBookName)
            
    Set FSO = New Scripting.FileSystemObject
    If FSO.GetFolder(ImportFilePath).Files.Count = 0 Then
       MsgBox "There are no files to import", vbInformation, APP_NAME
       Exit Sub
    End If

    Set VBModules = TargetBook.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each FileObj In FSO.GetFolder(ImportFilePath).Files
        Debug.Print FileObj.Name
        
        If (FSO.GetExtensionName(FileObj.Name) = "cls") Or _
            (FSO.GetExtensionName(FileObj.Name) = "frm") Or _
            (FSO.GetExtensionName(FileObj.Name) = "bas") And _
            FileObj.Name <> "ModProjectInOut.bas" Then
            VBModules.Import FileObj.Path
        End If
        
    Next FileObj
    Debug.Print "End of import"
End Sub
 
Public Sub RemoveAllModules()
    Dim ExportYN As Boolean
    Dim DlgOpen As FileDialog
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim EXPORTFILEPATH As String
    Dim ImportFileName As String
    Dim VBModule As VBIDE.VBComponent
   
    ''' NOTE: This workbook must be open in Excel.
    SourceBookName = ActiveWorkbook.Name
    Set SourceBook = Application.Workbooks(SourceBookName)
        
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        ''' remove it from the project if you want
        If VBModule.Type <> vbext_ct_Document Then SourceBook.VBProject.VBComponents.Remove VBModule
           
    Next VBModule
    
    Set DlgOpen = Nothing

End Sub

Public Sub ExportDBTables()
    Dim iFile As Integer
    Dim Fld As Field
    Dim FieldType As String
    Dim TableExport As TableDef
    Dim ExportFldr As String
        
    For Each TableExport In DB.TableDefs
        If Not (TableExport.Name Like "MSys*" Or TableExport.Name Like "~*") Then
            
            'debug.print TableExport.Name
            
            PrintFilePath = EXPORTFILEPATH & TableExport.Name & ".txt"
        
            iFile = FreeFile()
            
            Open PrintFilePath For Append As #iFile
            
            For Each Fld In TableExport.Fields
                Select Case Fld.Type
                    Case Is = 1
                        FieldType = "dbBoolean"
                    Case Is = 2
                        FieldType = "dbByte"
                    Case Is = 3
                        FieldType = "dbInteger"
                    Case Is = 4
                        FieldType = "dbLong"
                    Case Is = 5
                        FieldType = "dbCurrency"
                    Case Is = 6
                        FieldType = "dbSingle"
                    Case Is = 7
                        FieldType = "dbDouble"
                    Case Is = 8
                        FieldType = "dbDate"
                    Case Is = 9
                        FieldType = "dbBinary"
                    Case Is = 10
                        FieldType = "dbText"
                    Case Is = 11
                        FieldType = "dbLongBinary"
                    Case Is = 12
                        FieldType = "dbMemo"
                    Case Is = 15
                        FieldType = "dbGUID"
                    Case Is = 16
                        FieldType = "dbBigInt"
                    Case Is = 17
                        FieldType = "dbVarBinary"
                    Case Is = 18
                        FieldType = "dbChar"
                    Case Is = 19
                        FieldType = "dbNumeric"
                    Case Is = 20
                        FieldType = "dbDecimal"
                    Case Is = 21
                        FieldType = "dbFloat"
                    Case Is = 22
                        FieldType = "dbTime"
                    Case Is = 23
                        FieldType = "dbTimeStamp"
                    Case Is = 101
                        FieldType = "dbAttachment"
                    Case Is = 102
                        FieldType = "dbComplexByte"
                    Case Is = 103
                        FieldType = "dbComplexInteger"
                    Case Is = 104
                        FieldType = "dbComplexLong"
                    Case Is = 105
                        FieldType = "dbComplexSingle"
                    Case Is = 106
                        FieldType = "dbComplexDouble"
                    Case Is = 107
                        FieldType = "dbComplexGUID"
                    Case Is = 108
                        FieldType = "dbComplexDecimal"
                    Case Is = 109
                        FieldType = "dbComplexText"
                End Select
                
                Print #iFile, Fld.Name & ",  " & FieldType
            
            Next
                    
            Close #iFile
        End If
    Next
End Sub

Public Sub SetReferenceLibs()
    Dim Reference As Object
    
    On Error Resume Next
    
    For Each Reference In ThisWorkbook.VBProject.References
        With Reference
'            Debug.Print .Name
'            Debug.Print .Description
'            Debug.Print .Minor
'            Debug.Print .Major
'            Debug.Print .GUID
'            Debug.Print
        End With
    Next

    ' Visual Basic For Applications
    If Not ReferenceExists("{000204EF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{000204EF-0000-0000-C000-000000000046}", Major:=4, Minor:=1
    End If
    
    ' Microsoft Excel 14.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020813-0000-0000-C000-000000000046}", Major:=1, Minor:=7
    End If
    
    ' Microsoft Forms 2.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0D452EE1-E08F-101A-852E-02608C4D0BB4}", Major:=2, Minor:=0
    End If
    
    ' OLE Automation
    If Not ReferenceExists("{00020430-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020430-0000-0000-C000-000000000046}", Major:=2, Minor:=0
    End If
    
    ' Microsoft Office 14.0 Object Library
    If Not ReferenceExists("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", Major:=2, Minor:=5
    End If
    
    ' Microsoft Office 14.0 Access database engine Object Library
    If Not ReferenceExists("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", Major:=12, Minor:=0
    End If
    
    ' Microsoft Scripting Runtime
    If Not ReferenceExists("{420B2830-E718-11CF-893D-00A0C9054228}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0
    End If
    
    ' Microsoft Visual Basic for Applications Extensibility 5.3
    If Not ReferenceExists("{0002E157-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    End If
    
    ' Microsoft Outlook 14.0 Object Library
    If Not ReferenceExists("{00062FFF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00062FFF-0000-0000-C000-000000000046}", Major:=9, Minor:=4
    End If
End Sub

Public Function ReferenceExists(Ref As String) As Boolean
    Dim i As Integer
    
    With ThisWorkbook.VBProject.References
        For i = 1 To .Count
            If .Item(i).GUID = Ref Then
                ReferenceExists = True
                Exit Function
            End If
        Next
        ReferenceExists = False
    End With
End Function

Public Sub BuildProject()
    SetReferenceLibs
    ImportModules
    CopyShtCodeModule
End Sub
Public Sub CopyShtCodeModule()
    Dim SourceMod As VBIDE.VBComponent
    Dim DestMod As VBIDE.VBComponent
    Dim VBModule As VBIDE.VBComponent
    Dim VBCodeMod As VBIDE.CodeModule
    Dim i As Integer

    If ModuleExists("ThisWorkbook1") Then
        Set SourceMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook1")
        Set DestMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook")
    
        If DestMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
        End If
        
        If SourceMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
        End If
    End If
    
    For Each VBModule In ThisWorkbook.VBProject.VBComponents

        With VBModule

            Debug.Print VBModule.Name
            If Left(.Name, 3) = "Sht" And .Type <> vbext_ct_Document Then
                Set SourceMod = VBModule
                Debug.Print "Source: " & SourceMod.Name

                For Each DestMod In ThisWorkbook.VBProject.VBComponents
                    Debug.Print DestMod.Name
                    If Left(SourceMod.Name, Len(SourceMod.Name) - 1) = DestMod.Name Then
                        Debug.Print "Source: " & SourceMod.Name
                        Debug.Print " Dest: " & DestMod.Name

                        If SourceMod.CodeModule.CountOfLines > 0 Then
                            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
    
                            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
                        End If
                    End If
                Next
            End If
        End With
    Next

    For Each VBModule In ThisWorkbook.VBProject.VBComponents
        If Right(VBModule.Name, 1) = "1" And VBModule.Name <> "Sheet1" Then
            ThisWorkbook.VBProject.VBComponents.Remove VBModule
        End If
    Next VBModule



    Set SourceMod = Nothing
    Set DestMod = Nothing
    Set VBModule = Nothing
    Set VBCodeMod = Nothing
End Sub


Public Function ModuleExists(ModuleName As String) As Boolean
    Dim CodeModule As VBIDE.VBComponent
 
    For Each CodeModule In ThisWorkbook.VBProject.VBComponents
        If CodeModule.Name = ModuleName Then ModuleExists = True
    Next
End Function
