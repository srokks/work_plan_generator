VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ImageGroup As MSForms.image
Attribute ImageGroup.VB_VarHelpID = -1



Private Sub ImageGroup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
main.set_color Me.ImageGroup.BackColor
Unload color_form
End Sub

Private Sub ImageGroup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
gen_plan_dlgbox.choose_color_btn.BackColor = Me.ImageGroup.BackColor
End Sub
