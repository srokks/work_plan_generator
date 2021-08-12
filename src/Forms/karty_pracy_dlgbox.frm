VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} karty_pracy_dlgbox 
   Caption         =   "Generuj karty pracy"
   ClientHeight    =   5736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1980
   OleObjectBlob   =   "karty_pracy_dlgbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "karty_pracy_dlgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Checkboxes() As New checkbox_class

Private Sub gen_karty_Click()
karty_pracy.generuj
End Sub

Private Sub sheets_cmb_Change()
karty_pracy.update
End Sub

Private Sub UserForm_initialize()

For Each Sheet In Worksheets
    If UCase(Sheet.Name) = UCase("generator") Then
    Else
        sheets_cmb.AddItem (Sheet.Name)
    End If
Next

Dim checkBox As Control
square = 25
temp_licz = 1
For j = 0 To 7
    For i = 0 To 3
            Set checkBox = Frame1.Controls.Add("Forms.Checkbox.1")
                With checkBox
                    .Name = "day_btn_" + CStr(i)
                    .Caption = CStr(i)
                    .width = square
                    .height = square
                    .left = i * 20
                    .top = j * 20
                    .Caption = CStr(temp_licz)
                    .font.Name = "Cambria"
                    .font.Size = 6
                End With
                temp_licz = temp_licz + 1
            If temp_licz = 32 Then
                Exit For
            End If
    Next i
Next j


Dim ctl As Control
Dim checkBox_count As Integer: checkBox_count = 0

For Each ctl In Frame1.Controls
    If UCase(TypeName(ctl)) = UCase("checkbox") Then
        checkBox_count = checkBox_count + 1
        ReDim Preserve Checkboxes(1 To checkBox_count)
        Set Checkboxes(checkBox_count).CheckBoxGroup = ctl
        Checkboxes(checkBox_count).CheckBoxGroup_Enable
    End If
Next ctl

karty_pracy_dlgbox.company_name_txtbox.Text = Worksheets("Generator").Range("H16")

End Sub

