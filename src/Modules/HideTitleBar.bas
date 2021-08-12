Attribute VB_Name = "HideTitleBar"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long


    Public Declare PtrSafe Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Public Declare PtrSafe Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long
#Else
    Public Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long


    Public Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Public Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Public Declare Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long
#End If
Sub HideBar(frm As Object)

Dim Style As Long, Menu As Long, hWndForm As Long
hWndForm = FindWindow("ThunderDFrame", frm.Caption)
Style = GetWindowLong(hWndForm, &HFFF0)
Style = Style And Not &HC00000
SetWindowLong hWndForm, &HFFF0, Style
DrawMenuBar hWndForm

End Sub
