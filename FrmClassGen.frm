VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmClassGen 
   Caption         =   "UserForm1"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   OleObjectBlob   =   "FrmClassGen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmClassGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Const PRTE_OR_PUB As Integer = 0
Const VAR_NAME As Integer = 1
Const VAR_TYPE As Integer = 2
Const IS_PARENT As Integer = 3
Const IS_KEY As Integer = 4


Private Sub BtnAddMethod_Click()
    
    With LstMethods
        .AddItem
        .List(.ListCount - 1, 0) = CmoPublicPrivate2
        .List(.ListCount - 1, 1) = TxtMethodName
        .List(.ListCount - 1, 2) = CmoSubFunction
        .List(.ListCount - 1, 3) = TxtReturnType
        .List(.ListCount - 1, 4) = TxtDescription
        
    End With
    
    CmoPublicPrivate2 = ""
    TxtMethodName = ""
    CmoSubFunction = ""
    TxtReturnType = ""
    TxtDescription = ""
    
End Sub

Private Sub BtnAddVar_Click()
    Dim VariableStr As String
    
    With LstVariables
        .AddItem
        .List(.ListCount - 1, 0) = CmoPrivateOrPublic
        .List(.ListCount - 1, 1) = TxtName
        .List(.ListCount - 1, 2) = CmoType
        .List(.ListCount - 1, 3) = ChkParent
        .List(.ListCount - 1, 4) = ChkKey
        
    End With
    
    CmoPrivateOrPublic = ""
    TxtName = ""
    CmoType = ""
    ChkParent.Value = False
    ChkKey.Value = False '**
    
End Sub

Private Sub BtnDelete_Click()
        
    With LstVariables
        .RemoveItem .ListIndex
    End With
End Sub

Private Sub BtnDeleteMethod_Click()
    With LstMethods
        .RemoveItem .ListIndex
    End With

End Sub

Private Sub BtnGenerateClass_Click()
    
    
    Dim ClassFile
    Dim FileName As String
    Dim i As Integer
    Dim PrimaryKey As String
    
    ClassFile = FreeFile()
    
    FileName = "Cls" & TxtObjectName & " v0,0.cls"
    
    Open LIBRARY_FILE_PATH & FileName For Output As #ClassFile
    
    'Write header information
    Print #ClassFile, "VERSION 1.0 CLASS"
    Print #ClassFile, "BEGIN"
    Print #ClassFile, "    MultiUse = -1  'True"
    Print #ClassFile, "End"
    Print #ClassFile, "Attribute VB_Name = ""Cls" & TxtObjectName & """"
    Print #ClassFile, "Attribute VB_GlobalNameSpace = False"
    Print #ClassFile, "Attribute VB_Creatable = False"
    Print #ClassFile, "Attribute VB_PredeclaredId = False"
    Print #ClassFile, "Attribute VB_Exposed = False"
    Print #ClassFile, "'==============================================================="
    Print #ClassFile, "' Class Cls" & TxtObjectName
    Print #ClassFile, "' v0,0 - Initial Version"
    Print #ClassFile, "'---------------------------------------------------------------"
    Print #ClassFile, "' Date - " & Format(Now, "dd mmm yy")
    Print #ClassFile, "'==============================================================="
    Print #ClassFile, "' Methods"
    Print #ClassFile, "'---------------------------------------------------------------"
    
    With LstMethods
        For i = 0 To .ListCount - 1
            Print #ClassFile, "' " & .List(i, VAR_NAME) & " - " & .List(i, IS_KEY)
        Next
    End With
    Print #ClassFile, "'==============================================================="
    Print #ClassFile,
    Print #ClassFile, "Option Explicit"
    
    'write variables
    With LstVariables
        
        For i = 0 To .ListCount - 1
            If .List(i, PRTE_OR_PUB) = "Private" Then
                Print #ClassFile, "Private p" & .List(i, VAR_NAME) & " as " & .List(i, VAR_TYPE)
            Else
                Print #ClassFile, "Public " & .List(i, VAR_NAME) & " as " & .List(i, VAR_TYPE)
            End If
        Next
    End With
    
    If ChkChild Then
        Print #ClassFile, "Private pParent As Long"
    End If
    Print #ClassFile,
    Print #ClassFile, "'---------------------------------------------------------------"
    
    'Get /Set variables
    With LstVariables
        Dim TmpPublic As String
        Dim TmpVar As String
        Dim TmpType As String
        Dim IsClass As Boolean
        Dim ClassFlag As Boolean
        
        ClassFlag = False
        
        For i = 0 To .ListCount - 1
            
            TmpPublic = .List(i, PRTE_OR_PUB)
            TmpVar = .List(i, VAR_NAME)
            TmpType = .List(i, VAR_TYPE)
            If Left(TmpType, 3) = "Cls" Then
                IsClass = True
                ClassFlag = True
            Else
                IsClass = False
            End If
            
            'set primary key when found
            If .List(i, IS_KEY) = True Then
                PrimaryKey = .List(i, VAR_NAME)
            End If
                
            If TmpPublic = "Private" Then
                Print #ClassFile, "Public Property Get " & TmpVar & "() As " & TmpType
                
                If IsClass Then
                    Print #ClassFile, "    Set " & TmpVar & " = p" & TmpVar
                Else
                    Print #ClassFile, "    " & TmpVar & " = p" & TmpVar
                End If
                
                Print #ClassFile, "End Property"
                Print #ClassFile,
                Print #ClassFile, "Public Property Let " & TmpVar & "(ByVal vNewValue As " & TmpType & ")"
                
                If IsClass Then
                    Print #ClassFile, "    Set p" & TmpVar & " = vNewValue"
                Else
                    Print #ClassFile, "    p" & TmpVar & " = vNewValue"
                End If
                
                Print #ClassFile, "End Property"
                Print #ClassFile,
                Print #ClassFile, "'---------------------------------------------------------------"
            End If
        Next
    End With
        
    'Parent Get / Set
    If ChkChild Then
        Print #ClassFile, "Public Property Get Parent () As " & TxtChildOf
        Print #ClassFile, "    If pParent <> 0 Then"
        Print #ClassFile, "        Set Parent = GetParentFromPtr(pParent)"
        Print #ClassFile, "    End If"
        Print #ClassFile, "End Property"
        Print #ClassFile,
        Print #ClassFile, "Friend Function SetParent(ByVal Ptr As Long) As Boolean"
        Print #ClassFile, "    pParent = Ptr"
        Print #ClassFile, "End Function"
        Print #ClassFile,
    End If
    
    'Methods
    With LstMethods
        Dim TmpName As String
        Dim TmpSub As String
        Dim TmpReturn As String
        Dim TmpDesc As String
        Dim TmpParent As Boolean
        
        
        For i = 0 To .ListCount - 1
            TmpPublic = .List(i, PRTE_OR_PUB)
            TmpName = .List(i, VAR_NAME)
            TmpSub = .List(i, VAR_TYPE)
            TmpReturn = .List(i, IS_PARENT)
            TmpDesc = .List(i, IS_KEY)
            
            Print #ClassFile, "' ==============================================================="
            Print #ClassFile, "' Method " & TmpName
            Print #ClassFile, "' " & TmpDesc
            Print #ClassFile, "'---------------------------------------------------------------"
                        
            If TmpSub = "Sub" Then
                Print #ClassFile, "Public Sub " & TmpName
                Print #ClassFile,
                Print #ClassFile, "End Sub"
            Else
                Print #ClassFile, "Public Function " & TmpName & "() As " & TmpReturn
                Print #ClassFile,
                Print #ClassFile, "End Function"
            End If
            Print #ClassFile,
        Next
    End With
    
    'Get Parent Method
    If ChkChild Then
        Print #ClassFile, "' ==============================================================="
        Print #ClassFile, "' Method GetParentFromPtr"
        Print #ClassFile, "' Private routine to copy memory address of parent class"
        Print #ClassFile, "' ---------------------------------------------------------------"
        Print #ClassFile, "Private Function GetParentFromPtr(ByVal Ptr As Long) As " & TxtChildOf
        Print #ClassFile, "    Dim tmp As " & TxtChildOf
        Print #ClassFile,
        Print #ClassFile, "    CopyMemory tmp, Ptr, 4"
        Print #ClassFile, "    Set GetParentFromPtr = tmp"
        Print #ClassFile, "    CopyMemory tmp, 0&, 4"
        Print #ClassFile, "End Function"
    End If
    
    'initialise
    If ClassFlag Then
        Print #ClassFile,
        Print #ClassFile, "' ==============================================================="
        Print #ClassFile, "Private Sub Class_Initialize()"
                
        With LstVariables
            IsClass = False
            For i = 0 To .ListCount - 1
                TmpVar = .List(i, VAR_NAME)
                TmpType = .List(i, VAR_TYPE)
                If Left(TmpType, 3) = "Cls" Then IsClass = True Else IsClass = False
                
                If IsClass Then
                    Print #ClassFile, "    Set p" & TmpVar & " = New " & TmpType
                End If
            Next
            
            Print #ClassFile,
            
            For i = 0 To .ListCount - 1
                TmpPublic = .List(i, PRTE_OR_PUB)
                TmpVar = .List(i, VAR_NAME)
                TmpType = .List(i, VAR_TYPE)
                If .List(i, IS_PARENT) = True Then TmpParent = .List(i, IS_PARENT)
                
                If TmpParent = True Then
                    If TmpPublic = "Private" Then
                        Print #ClassFile, "    p" & TmpVar & ".SetParent ObjPtr(Me)"
                    Else
                        Print #ClassFile, "    " & TmpVar & ".SetParent ObjPtr(Me)"
                    End If
                End If
            Next
        End With
        
        Print #ClassFile, "End Sub"
        Print #ClassFile,
        Print #ClassFile, "'---------------------------------------------------------------"
        Print #ClassFile, "Private Sub Class_Terminate()"
    
        'Terminate
        
        With LstVariables
            IsClass = False
            For i = 0 To .ListCount - 1
                TmpVar = .List(i, VAR_NAME)
                TmpType = .List(i, VAR_TYPE)
                If Left(TmpType, 3) = "Cls" Then IsClass = True Else IsClass = False
                
                If IsClass Then
                    Print #ClassFile, "    Set p" & TmpVar & " = Nothing"
                End If
            Next
            
            Print #ClassFile,
            
            For i = 0 To .ListCount - 1
                TmpPublic = .List(i, PRTE_OR_PUB)
                TmpVar = .List(i, VAR_NAME)
                TmpType = .List(i, VAR_TYPE)
                If .List(i, IS_PARENT) = True Then TmpParent = .List(i, IS_PARENT)
                
                If TmpParent = True Then
                    If TmpPublic = "Private" Then
                        Print #ClassFile, "    p" & TmpVar & ".SetParent 0"
                    Else
                        Print #ClassFile, "    " & TmpVar & ".SetParent ObjPtr(Me)"
                    End If
                End If
            Next
        End With
        Print #ClassFile, "End Sub"
        Print #ClassFile,
        Print #ClassFile, "'---------------------------------------------------------------"
    End If
        
    
    Close #ClassFile
    
End Sub




Private Sub BtnGenerateDBTable_Click()
    Dim TableDef As DAO.TableDef
    Dim TmpVarName As String
    Dim TmpVarType As String
    Dim i As Integer
    
    Dim Fld As DAO.Field
    
    Initialise
    
    Set TableDef = DB.CreateTableDef("Tbl" & TxtObjectName)
    
    With TableDef
        For i = 0 To LstVariables.ListCount - 1
                        
            TmpVarName = LstVariables.List(i, 1)
            Select Case LstVariables.List(i, 2)
                Case Is = "String"
                    Set Fld = .CreateField(TmpVarName, dbText)
                    .Fields.Append Fld
                Case Is = "Integer"
                    Set Fld = .CreateField(TmpVarName, dbInteger)
                    .Fields.Append Fld
                Case Is = "Date"
                    Set Fld = .CreateField(TmpVarName, dbDate)
                    .Fields.Append Fld
                Case Is = "Boolean"
                    Set Fld = .CreateField(TmpVarName, dbBoolean)
                    .Fields.Append Fld
            End Select
                        
        Next
    End With
    DB.TableDefs.Append TableDef
    
    Set TableDef = Nothing
    Set Fld = Nothing
    
End Sub

Private Sub ChkDBClass_Click()
    If Me.Enabled = False Then LstMethods.Clear
End Sub

Private Sub UserForm_Activate()
    With CmoPrivateOrPublic
        .Clear
        .AddItem "Private"
        .AddItem "Public"
    End With
    
    With CmoPublicPrivate2
        .Clear
        .AddItem "Private"
        .AddItem "Public"
    End With
    
    With CmoSubFunction
        .Clear
        .AddItem "Sub"
        .AddItem "Function"
    End With
    
    With CmoType
        .Clear
        .AddItem "String"
        .AddItem "Integer"
        .AddItem "Date"
        .AddItem "Boolean"
    End With
    

    With LstVariables
        .Clear
        .AddItem
        .List(0, 0) = "Private"
        .List(0, 1) = "Deleted"
        .List(0, 2) = "Date"
        
    End With
    
End Sub


