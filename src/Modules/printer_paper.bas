Attribute VB_Name = "printer_paper"
'Written: June 14, 2010
'Author:  Leith Ross
'Summary: Lists the supported paper sizes for the default printer in a message box.

Private Const DC_PAPERNAMES = &H10
#If VBA7 Then
    Private Declare PtrSafe Function DeviceCapabilities _
      Lib "winspool.drv" _
        Alias "DeviceCapabilitiesA" _
          (ByVal lpDeviceName As String, _
           ByVal lpPort As String, _
           ByVal iIndex As Long, _
           ByRef lpOutput As Any, _
           ByRef lpDevMode As Any) _
        As Long
    
    Private Declare PtrSafe Function StrLen _
      Lib "kernel32.dll" _
        Alias "lstrlenA" _
          (ByVal lpString As String) _
        As Long
#Else
    Private Declare Function DeviceCapabilities _
      Lib "winspool.drv" _
        Alias "DeviceCapabilitiesA" _
          (ByVal lpDeviceName As String, _
           ByVal lpPort As String, _
           ByVal iIndex As Long, _
           ByRef lpOutput As Any, _
           ByRef lpDevMode As Any) _
        As Long
    
    Private Declare Function StrLen _
      Lib "kernel32.dll" _
        Alias "lstrlenA" _
          (ByVal lpString As String) _
        As Long
#End If
Sub ListPaperSizes()

  Dim AllNames As String
  Dim i As Long
  Dim Msg As String
  Dim PD As Variant
  Dim Ret As Long
  Dim PaperSizes() As Byte
  Dim PaperSize As String

   'Retrieve the number of available paper names
    PD = Split(Application.ActivePrinter, " na ")
    Ret = DeviceCapabilities(PD(0), PD(1), DC_PAPERNAMES, ByVal 0&, ByVal 0&)

   'resize the array
    ReDim PaperSizes(0 To Ret * 64) As Byte

   'retrieve all the available paper names
    Call DeviceCapabilities(PD(0), PD(1), DC_PAPERNAMES, PaperSizes(0), ByVal 0&)

   'convert the retrieved byte array to an ANSI string
    AllNames = StrConv(PaperSizes, vbUnicode)

     'loop through the string and search for the names of the papers
      For i = 1 To Len(AllNames) Step 64
        PaperSize = Mid(AllNames, i, 64)
        PaperSize = left(PaperSize, StrLen(PaperSize))
        Msg = Msg & PaperSize & vbCrLf
      Next i

    MsgBox "Supported Paper Sizes:" & vbCrLf & vbCrLf & Msg, vbOKOnly, PD(0)
    
End Sub

Function printer_a3_comp() As Boolean
Dim AllNames As String
Dim i As Long
Dim Msg As String
Dim PD As Variant
Dim Ret As Long
Dim PaperSizes() As Byte
Dim PaperSize As String
'Retrieve the number of available paper names
PD = Split(Application.ActivePrinter, " na ")
Ret = DeviceCapabilities(PD(0), PD(1), DC_PAPERNAMES, ByVal 0&, ByVal 0&)

'resize the array
 ReDim PaperSizes(0 To Ret * 64) As Byte

'retrieve all the available paper names
 Call DeviceCapabilities(PD(0), PD(1), DC_PAPERNAMES, PaperSizes(0), ByVal 0&)

'convert the retrieved byte array to an ANSI string
 AllNames = StrConv(PaperSizes, vbUnicode)

'loop through the string and search for the names of the papers
 For i = 1 To Len(AllNames) Step 64
   PaperSize = Mid(AllNames, i, 64)
   PaperSize = left(PaperSize, StrLen(PaperSize))
   If PaperSize = "A3" Then
        check_printer_paper = True
   Else
        check_printer_paper = False
   End If
 Next i

End Function
