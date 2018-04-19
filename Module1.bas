Attribute VB_Name = "Module1"
Sub CreateUserForm()
Dim myForm As Object
Dim NewFrame As MSForms.Frame
Dim NewButton As MSForms.CommandButton
'Dim NewComboBox As MSForms.ComboBox
Dim NewListBox As MSForms.ListBox
'Dim NewTextBox As MSForms.TextBox
'Dim NewLabel As MSForms.Label
'Dim NewOptionButton As MSForms.OptionButton
'Dim NewCheckBox As MSForms.CheckBox
Dim X As Integer
Dim Line As Integer

Set myForm = ThisWorkbook.VBProject.VBComponents.Add(3)

'Create the User Form
With myForm
    .Properties("Caption") = "New Form"
    .Properties("Width") = 300
    .Properties("Height") = 270
End With

'Create ListBox
Set NewListBox = myForm.Designer.Controls.Add("Forms.listbox.1")
With NewListBox
    .Name = "lst_1"
    .Top = 10
    .Left = 10
    .Width = 150
    .Height = 230
    .Font.Size = 8
    .Font.Name = "Tahoma"
    .BorderStyle = fmBorderStyleOpaque
    .SpecialEffect = fmSpecialEffectSunken
End With

'Create CommandButton Create
Set NewButton = myForm.Designer.Controls.Add("Forms.commandbutton.1")
With NewButton
    .Name = "cmd_1"
    .Caption = "clickMe"
    .Accelerator = "M"
    .Top = 10
    .Left = 200
    .Width = 66
    .Height = 20
    .Font.Size = 8
    .Font.Name = "Tahoma"
    .BackStyle = fmBackStyleOpaque
End With

'add code for listBox
lstBoxData = "Data 1,Data 2,Data 3,Data 4"
myForm.CodeModule.InsertLines 1, "Private Sub UserForm_Initialize()"
myForm.CodeModule.InsertLines 2, "   me.lst_1.addItem ""Data 1"" "
myForm.CodeModule.InsertLines 3, "   me.lst_1.addItem ""Data 2"" "
myForm.CodeModule.InsertLines 4, "   me.lst_1.addItem ""Data 3"" "
myForm.CodeModule.InsertLines 5, "End Sub"

'add code for Comand Button
myForm.CodeModule.InsertLines 6, "Private Sub cmd_1_Click()"
myForm.CodeModule.InsertLines 7, "   If me.lst_1.text <>"""" Then"
myForm.CodeModule.InsertLines 8, "      msgbox (""You selected item: "" & me.lst_1.text )"
myForm.CodeModule.InsertLines 9, "   End If"
myForm.CodeModule.InsertLines 10, "End Sub"
'Show the form
VBA.UserForms.Add(myForm.Name).Show

'Delete the form (Optional)
'ThisWorkbook.VBProject.VBComponents.Remove myForm
End Sub
