Attribute VB_Name = "pointer"
Option Explicit
' these are special function to get device specific things

Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
 
Declare PtrSafe Function GetDeviceCaps Lib "Gdi32" (ByVal hDC As Long, _
     ByVal nIndex As Long) As Long
 
Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
     ByVal hDC As Long) As Long
 
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
 
' we need to be able to find cursor position where mouse was clicked
Public Type tCursor
    left As Long
    top As Long
End Type
 
Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As tCursor) As Long
 
Public Function pointsPerPixelX() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function
 
Public Function pointsPerPixelY() As Double
    Dim hDC As Long
    hDC = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
End Function
 
Public Function WhereIsTheMouseAt() As tCursor
    Dim mPos As tCursor
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function
Public Function convertMouseToForm() As tCursor
    Dim mPos As tCursor
    mPos = WhereIsTheMouseAt
    mPos.left = pointsPerPixelY * mPos.left
    mPos.top = pointsPerPixelX * mPos.top
    convertMouseToForm = mPos
End Function

