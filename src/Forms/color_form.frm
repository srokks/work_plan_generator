VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} color_form 
   Caption         =   "Color choose"
   ClientHeight    =   1680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   1788
   OleObjectBlob   =   "color_form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "color_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Images() As New ImageClass
Private workws As Worksheet
Private Sub main_frame_Initialize()
Dim imagecout As Integer: imagecout = 0
Dim ctl As Control
source_sheet = "Generator"
Set workws = Worksheets(source_sheet)
Dim image As Control
Dim row_pos As Integer
square = 12
For i = 1 To 3
    For j = 1 To 9
        Set image = main_frame.Controls.Add("Forms.Image.1")
        With image
            .Name = "col" + CStr(i)
            .height = square
            .width = square
            .top = (i - 1) * square
            .left = (j - 1) * square
            .BackColor = workws.Range(Cells(i, j + 12), Cells(i, j + 12)).Interior.Color
            .BorderStyle = 0
        End With
        
'        If i = 1 Then
'            With image
'                .Name = "col" + CStr(i)
'                .height = 10
'                .top = j * 10
'                .left = 0
'                .width = 10
'                .BackColor = workWS.Range(Cells(i, j + 12), Cells(i, j + 12)).Interior.Color
'            End With
'        Else
'            With image
'                .Name = "col" + CStr(i)
'                .height = 10
'                .top = j * 10
'                .left = (i - 1) * 10
'                .width = 10
'                .BackColor = workWS.Range(Cells(i, j + 12), Cells(i, j + 12)).Interior.Color
'            End With
'        End If

    Next j
Next i

For Each ctl In Me.Controls
    If TypeName(ctl) = "Image" Then
        imagecout = imagecout + 1
        ReDim Preserve Images(1 To imagecout)
        Set Images(imagecout).ImageGroup = ctl
    End If
Next ctl

End Sub

Private Sub UserForm_initialize()
transparent.UserformTransparent color_form, 255

'ustawienie pozycji okientka
Dim pos As pointer.tCursor
pos = pointer.WhereIsTheMouseAt
Me.top = pointer.pointsPerPixelX * pos.top
Me.left = pointer.pointsPerPixelY * pos.left
Me.top = Me.top + 10
Me.left = Me.left + 10
HideTitleBar.HideBar Me
Me.height = 43
Me.width = 102

main_frame_Initialize



End Sub

