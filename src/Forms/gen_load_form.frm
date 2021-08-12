VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gen_load_form 
   Caption         =   "UserForm2"
   ClientHeight    =   1488
   ClientLeft      =   132
   ClientTop       =   552
   ClientWidth     =   4740
   OleObjectBlob   =   "gen_load_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gen_load_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_initialize()
HideTitleBar.HideBar Me
Me.height = Me.height - 25
Me.width = Me.width - 10
'transparent.UserformTransparent Me, 255
End Sub

