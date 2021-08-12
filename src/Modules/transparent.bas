Attribute VB_Name = "transparent"
Option Explicit

Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                     Alias "GetWindowLongA" _
                    (ByVal hWnd As LongPtr, _
                     ByVal nIndex As Long) As LongPtr

Private Declare PtrSafe Function SetWindowLong Lib "user32" _
                     Alias "SetWindowLongA" _
                    (ByVal hWnd As LongPtr, _
                     ByVal nIndex As Long, _
                     ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
ByVal hWnd As LongPtr, _
ByVal crey As Byte, _
ByVal bAlpha As Byte, _
ByVal dwFlags As Long) As LongPtr

Private Const GWL_EXSTYLE       As Long = (-20)
Private Const LWA_COLORKEY      As Long = &H1
Private Const LWA_ALPHA         As Long = &H2 'H2
Private Const WS_EX_LAYERED     As Long = &HFF0000

Public Declare PtrSafe Function FindWindowA Lib "user32" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'
'
'   *- TRANSPARENCE : SUPPR COULEUR / FORM ALPHA (auteur inconnu) -*
'   =============================================================
Public Function WndSetOpacity(ByVal hWnd As LongPtr, Optional ByVal crKey As Long = vbBlack, Optional ByVal Alpha As Byte = 255, Optional ByVal ByAlpha As Boolean = True) As Boolean
' Return : True si il n'y a pas eu d'erreur.
' hWnd   : hWnd de la fenetre a rendre transparente
' crKey  : Couleur a rendre transparente si ByAlpha=False (utiliser soit les constantes vb:vbWhite ou en hexa:&HFFFFFF)
' Alpha  : 0-255 0=transparent 255=Opaque si ByAlpha=true (défaut)
On Error GoTo Lbl_Exit

Dim ExStyle As LongPtr
ExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
If ExStyle <> (ExStyle Or WS_EX_LAYERED) Then
    ExStyle = (ExStyle Or WS_EX_LAYERED)
    Call SetWindowLong(hWnd, GWL_EXSTYLE, ExStyle)
End If
WndSetOpacity = (SetLayeredWindowAttributes(hWnd, crKey, Alpha, IIf(ByAlpha, LWA_COLORKEY Or LWA_ALPHA, LWA_COLORKEY)) <> 0)

Lbl_Exit:
On Error GoTo 0
If Not Err.Number = 0 Then Err.Clear
End Function

Public Sub UserformTransparent(ByRef uf As Object, TransparenceControls As Integer)
'uf as MSForms.UserForm won't work !!!!
Dim B As Boolean
Dim lHwnd As LongPtr
On Error GoTo 0
'- Recherche du handle de la fenetre par son Caption
lHwnd = FindWindowA(vbNullString, uf.Caption)
If lHwnd = 0 Then
    MsgBox "Handle de " & uf.Caption & " Introuvable", vbCritical
    Exit Sub
End If
'If d And F Then
    B = WndSetOpacity(lHwnd, uf.BackColor, TransparenceControls, True)
'ElseIf d Then
'    'B = WndSetOpacity(M.hwnd, , 255, True)
'    B = WndSetOpacity(lHwnd, , TransparenceControls, True)
'Else
'    B = WndSetOpacity(lHwnd, , 255, True)
'End If
End Sub


Public Sub ActiveTransparence(stCaption As String, d As Boolean, F As Boolean, Couleur As Long, Transparence As Integer)
Dim B As Boolean
Dim lHwnd As Long
'- Recherche du handle de la fenetre par son Caption
lHwnd = FindWindowA(vbNullString, stCaption)
If lHwnd = 0 Then
    MsgBox "Handle de " & stCaption & " Introuvable", vbCritical
    Exit Sub
End If
If d And F Then
    B = WndSetOpacity(lHwnd, Couleur, Transparence, True)
ElseIf d Then
    'B = WndSetOpacity(M.hwnd, , 255, True)
    B = WndSetOpacity(lHwnd, , Transparence, True)
Else
    B = WndSetOpacity(lHwnd, , 255, True)
End If
End Sub

