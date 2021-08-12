VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gen_plan_dlgbox 
   Caption         =   "Generuj pusty grafik"
   ClientHeight    =   4044
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2172
   OleObjectBlob   =   "gen_plan_dlgbox.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gen_plan_dlgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub choose_color_btn_Click()
    Me.Enabled = False
    color_form.Show
    Me.Enabled = True
End Sub


Private Sub choose_month_cmbbox_Change()
update_sheet
End Sub

Private Sub choose_year_cmbbox_Change()
update_sheet
End Sub

Private Sub gen_btn_Click()
main.generuj
End Sub

Private Sub prev_data_btn_Click()
main.prev_data
End Sub


Private Sub sheet_name_txtbox_Change()
user_sheet_name.Enabled = True
Me.gen_btn.Enabled = False
End Sub

Private Sub user_sheet_name_Click()
main.update_sheet_name
user_sheet_name.Enabled = False
gen_btn.Enabled = True
End Sub

Public Sub UserForm_initialize()
main.init_form
copyright_label.Caption = "©Copyright 2021,Jaros³aw Sroka" + Chr(10) + "All rights reserved."
End Sub
