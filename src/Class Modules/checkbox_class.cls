VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "checkbox_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents CheckBoxGroup As MSForms.checkBox
Attribute CheckBoxGroup.VB_VarHelpID = -1





Public Sub CheckBoxGroup_Click()
karty_pracy.update
End Sub

Public Sub CheckBoxGroup_Enable()
CheckBoxGroup.Enabled = True
End Sub

