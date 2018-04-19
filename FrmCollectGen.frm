VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCollectGen 
   Caption         =   "UserForm1"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   OleObjectBlob   =   "FrmCollectGen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCollectGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' Form FrmCollectGen
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================


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
        
    End With
    
    CmoPrivateOrPublic = ""
    TxtName = ""
    CmoType = ""
    ChkParent.Value = False
    
End Sub

Private Sub BtnDeleteMethod_Click()
    With LstMethods
        .RemoveItem .ListIndex
    End With

End Sub

Private Sub BtnDeleteVar_Click()
    With LstVariables
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
    
    Open ShtSettings.Range("FPath") & FileName For Output As #ClassFile
    
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
    Print #ClassFile, "Private p"; TxtObjectName & " As Collection"

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
    End If
    
    'Methods
    With LstMethods
        Dim TmpName As String
        Dim TmpSub As String
        Dim TmpReturn As String
        Dim TmpDesc As String
        Dim TmpParent As Boolean
        
        Print #ClassFile,
        Print #ClassFile, "Public Function NewEnum() as IUnknown"
        Print #ClassFile, "Attribute NewEnum.VB_UserMemId =-4"
        Print #ClassFile, "    Set newEnum = mcolcells.[_newenum]"
        Print #ClassFile, "End function"
        Print #ClassFile,
        
        For i = 0 To .ListCount - 1
            TmpPublic = .List(i, PRTE_OR_PUB)
            TmpName = .List(i, VAR_NAME)
            TmpSub = .List(i, VAR_TYPE)
            TmpDesc = .List(i, IS_KEY)
            
            Print #ClassFile, "' ==============================================================="
            Print #ClassFile, "' Method " & TmpName
            Print #ClassFile, "' " & TmpDesc
            Print #ClassFile, "'---------------------------------------------------------------"
            Print #ClassFile,
            Print #ClassFile, GetMethodScript(i)
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
    Print #ClassFile,
    Print #ClassFile, "' ==============================================================="
    Print #ClassFile, "Private Sub Class_Initialize()"
    Print #ClassFile, "    Set p" & TxtObjectName & " = New Collection"
    
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
            TmpParent = .List(i, IS_PARENT)
            
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
    
    Print #ClassFile, "    Set p" & TxtObjectName & " = nothing"
    
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
            TmpParent = .List(i, IS_PARENT)
            
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

        
    
    Close #ClassFile
    Hide
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
                Case Is = "Integer"
                    Set Fld = .CreateField(TmpVarName, dbInteger)
                Case Is = "Date"
                    Set Fld = .CreateField(TmpVarName, dbDate)
                Case Is = "Boolean"
                    Set Fld = .CreateField(TmpVarName, dbBoolean)
            End Select
                        
            .Fields.Append Fld
        Next
    End With
    DB.TableDefs.Append TableDef
    
    Set TableDef = Nothing
    Set Fld = Nothing
    
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
    
    With LstMethods
        .Clear
        .AddItem
        .List(0, 0) = "Public"
        .List(0, 1) = "FindItem"
        .List(0, 2) = "Function"
        .List(0, 4) = "Finds Item from index"
        .AddItem
        .List(1, 0) = "Public"
        .List(1, 1) = "AddItem"
        .List(1, 2) = "Sub"
        .List(1, 4) = "Add item to Collection"
        .AddItem
        .List(2, 0) = "Public"
        .List(2, 1) = "RemoveItem"
        .List(2, 2) = "Sub"
        .List(2, 4) = "Deletes item from collection"
        .AddItem
        .List(3, 0) = "Public"
        .List(3, 1) = "ItemCount"
        .List(3, 2) = "Function"
        .List(3, 4) = "Returns number of items in collection"
        .AddItem
        .List(4, 0) = "Public"
        .List(4, 1) = "GetCollection"
        .List(4, 2) = "Sub"
        .List(4, 4) = "Gets all items in collection"
        .AddItem
        .List(5, 0) = "Public"
        .List(5, 1) = "DeleteCollection"
        .List(5, 2) = "Sub"
        .List(5, 4) = "Deletes all items in collection"
        
    
    End With

End Sub


Private Function GetMethodScript(i As Integer) As String
    Dim Txt As String
    Dim CollectionType As String
    Dim X As Integer
    
    CollectionType = Left(TxtObjectName, Len(TxtObjectName) - 1)
        
    Select Case i
        Case 0
            Txt = "Public Function FindItem(" & TxtIndex & " As Variant) As Cls" & CollectionType _
                    & Chr(13) & " Attribute item.VB_UserMemId = 0" _
                    & Chr(13) & "    On Error Resume Next" & Chr(13) & "    Set FindItem = p" & TxtObjectName & ".Item(" & TxtIndex & ")" _
                    & Chr(13) & "End Function"

        Case 1
            Txt = "Public Sub AddItem(" & CollectionType & " As Cls" & CollectionType & ")" _
                    & Chr(13) & "    " & CollectionType & ".SetParent ObjPtr(me)" _
                    & Chr(13) & "    p" & TxtObjectName & ".Add " & CollectionType & ", Key:=CStr(" & CollectionType & "." & TxtIndex & ")" _
                    & Chr(13) & "End Sub"

        Case 2
            Txt = "Public Sub RemoveItem(" & TxtIndex & " As Variant)" _
                    & Chr(13) & "    " & "p" & TxtObjectName & ".Remove " & TxtIndex _
                    & Chr(13) & "End Sub"

        Case 3
            Txt = "Public Function Count() As Integer" _
                    & Chr(13) & "    " & "Count = p" & TxtObjectName & ".Count" _
                    & Chr(13) & "End Function"
        
        Case 4
            Txt = "Public Sub GetCollection()" _
                & Chr(13) & "    Dim Rst" & CollectionType & " As Recordset" _
                & Chr(13) & "    Dim " & CollectionType & " As Cls" & CollectionType _
                & Chr(13) & "    Dim i As Integer" _
                & Chr(13) & "" _
                & Chr(13) & "    Set Rst" & CollectionType & " = ModDatabase.SQLQuery(""SELECT * FROM Tbl" & CollectionType & " WHERE Deleted IS NULL"")" _
                & Chr(13) & "    With Rst" & CollectionType _
                & Chr(13) & "        .MoveLast" _
                & Chr(13) & "        .MoveFirst" _
                & Chr(13) & "        For i = 1 to .recordcount" _
                & Chr(13) & "            Set " & CollectionType & " = New Cls" & CollectionType _
                & Chr(13) & "            " & CollectionType & "." & TxtIndex & " = !" & TxtIndex _
                & Chr(13) & "            " & CollectionType & ".DBGet" _
                & Chr(13) & "            Me.AddItem " & CollectionType _
                & Chr(13) & "            .Movenext" _
                & Chr(13) & "        Next" _
                & Chr(13) & "    End with" _
                & Chr(13) & "End Sub" _

        Case 5
            Txt = "Public Sub DeleteCollection()" _
                  & Chr(13) & "    Dim " & CollectionType & " As Cls" & CollectionType _
                    & Chr(13) & "    For Each " & CollectionType & " In p" & TxtObjectName _
                    & Chr(13) & "        p" & TxtObjectName & ".Remove cstr(" & CollectionType & "." & TxtIndex & ")" _
                    & Chr(13) & "        " & CollectionType & ".DBDelete" _
                    & Chr(13) & "    Next" _
                    & Chr(13) & "End Sub"
   
    End Select
    GetMethodScript = Txt
End Function
