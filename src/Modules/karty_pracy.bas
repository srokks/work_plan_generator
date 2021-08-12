Attribute VB_Name = "karty_pracy"
Private days_count As Integer
Private sourceWB As Workbook
Private sourceWS As Worksheet
Private workws As Worksheet
Private sheet_name, source_name As String
Private saint_days() As Integer
Private kartyWB As Workbook
Public source_path, cur_year, cur_month, comp_name As String
Public hours As Integer
Private width, height  As Integer
Public Start As Variant
Public Sub update()
If karty_pracy_dlgbox.sheets_cmb.Value <> "" Then
    source_name = karty_pracy_dlgbox.sheets_cmb.Value
    Set sourceWS = Worksheets(source_name)
    karty_pracy_dlgbox.gen_karty.Enabled = True
End If
Dim ctl As Control
Dim licznik As Integer: licznik = 0
For Each ctl In karty_pracy_dlgbox.Frame1.Controls
    If ctl.Value Then
        licznik = licznik + 1
        ReDim Preserve saint_days(1 To licznik)
        saint_days(licznik) = ctl.Caption
    End If
    If licznik = 0 Then
        licznik = licznik + 1
        ReDim Preserve saint_days(1 To licznik)
        saint_days(licznik) = 0
    End If
Next
source_path = ActiveWorkbook.Path
comp_name = karty_pracy_dlgbox.company_name_txtbox
End Sub

Private Sub fill_SpecialDays(ByVal workws As Variant)
Dim specialPhrases As Variant
specialPhrases = Array("wn", "w5", "ws", "l4", "nn", "nu")
If IsNumeric(UBound(saint_days)) = False Then MsgBox "a is empty!"

For i = 3 To width
   
    If workws.Range(Cells(3, i), Cells(3, i)) = "sb" Or workws.Range(Cells(3, i), Cells(3, i)) = "nd" Or IsInArray(Day(workws.Range(Cells(4, i), Cells(4, i))), saint_days) Then

        For j = 6 To height Step 2
            If workws.Range(Cells(j, i), Cells(j, i)) = "" Then
                Select Case workws.Range(Cells(3, i), Cells(3, i))
                    Case "sb"
                        workws.Range(Cells(j, i), Cells(j, i)) = "w5"
                    Case "nd"
                        workws.Range(Cells(j, i), Cells(j, i)) = "wn"
                End Select
                If IsInArray(Day(workws.Range(Cells(4, i), Cells(4, i))), saint_days) Then
                    workws.Range(Cells(j, i), Cells(j, i)) = "ws"
                End If
                
            ElseIf workws.Range(Cells(j, i), Cells(j, i)) <> "" And Not (IsInArray(workws.Range(Cells(j, i), Cells(j, i)), specialPhrases)) Then
                workws.Range(Cells(j - 1, i), Cells(j, i + 1)).Select
                With Selection
                    If workws.Range(Cells(3, i), Cells(3, i)) = "sb" Then
                        .Interior.Color = 11389944
                    ElseIf workws.Range(Cells(3, i), Cells(3, i)) = "nd" Then
                        .Interior.Color = 15652797
                    ElseIf IsInArray(Day(workws.Range(Cells(4, i), Cells(4, i))), saint_days) Then
                        .Interior.Color = 7434751
                    End If
                        
                End With
            End If
        Next
    End If
Next
End Sub

Private Sub create_new_sheet(ByVal file_name As String)
sourceWS.Copy
Set kartyWB = ActiveWorkbook
ActiveWorkbook.SaveAs Filename:=source_path + "\" + file_name + ".xlsx"
End Sub


Private Sub transformPlan()
'ci¹g³y b³¹d: out of script range - rozwi¹zany przez nie dos³owne przypisanie nazwy, mo¿na zmieniæ na nazwa+xlxs

Set workws = Worksheets(source_name)
set_dimensions
'wymazuje wykrzykniki
For i = 1 To height
    For j = 1 To width
        If workws.Range(Cells(i, j), Cells(i, j)) = "!" Then
            workws.Range(Cells(i, j), Cells(i, j)).ClearContents
        End If
    Next
Next
workws.Columns(1).ColumnWidth = 2.22
If workws.Range("A1") <> "HARMONOGRAM PRACY" Then
    workws.Rows("1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    workws.Range(Cells(1, 1), Cells(1, width - 2)).Merge
    workws.Range(Cells(1, 1), Cells(1, width - 2)) = "HARMONOGRAM PRACY"
    workws.Range(Cells(1, 1), Cells(1, width)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .Borders.Color = vbBlack
        .Borders.Weight = xlThin
        .font.Name = "Cambria"
        .font.Size = 9
        .font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = workws.Range("B3").Interior.Color
    End With
End If
cur_month = workws.Range("B3")
cur_year = CInt(Year(workws.Range("C4")))
hours = workws.Range("A3")
workws.Range("B4") = workws.Range("B3") + "; " + CStr(workws.Range("A3"))
workws.Range("B3") = "miesi¹c; norma"
workws.Range("A3") = ""
'czyszczenie
workws.Range(Cells(3, 3), Cells(height, width - 2)).Interior.Pattern = xlNone
workws.Name = "HARMONOGRAM PRACY"
fill_SpecialDays workws
transformationPolish
End Sub

Private Sub set_dimensions()
height = 5
width = 3
Do While workws.Range(Cells(4, width), Cells(4, width)).Value <> ""
    width = width + 2
Loop
Do While workws.Range(Cells(height, 2), Cells(height, 2)).Value <> ""
    height = height + 2
Loop
End Sub

Private Sub transformationPolish()
workws.Range("a1").UnMerge
workws.Range(Cells(1, 1), Cells(1, width - 2)).Merge
workws.Range(Cells(1, width + 1), Cells(height, width + 10)).Clear
workws.Range(Cells(height, 1), Cells(height + 5, width + 5)).Clear
Application.PrintCommunication = True
With workws.PageSetup
        .LeftHeader = ""
        .RightHeader = ""
    End With
Application.PrintCommunication = False
workws.Range(Cells(1, 1), Cells(height + 10, width + 10)).ClearComments
End Sub

Public Sub generuj()
'Prod source_path
'source_path = ActiveWorkbook.Path

Unload karty_pracy_dlgbox
Start = Timer
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
gen_load_form.progres_label.width = 0
gen_load_form.Show
'1 i 2 .Skopiowanie Ÿród³a do nowego arkusza i go Utworzenie  - Karty pracy - Nazwa Ÿród³a
create_new_sheet "Karty Pracy - " + source_name
gen_load_form.progres_label.width = 0
gen_load_form.Show

main.progress_bar 19, "create_new_sheet"
'3.Transformacja arksza pod karty pracy
transformPlan
main.progress_bar 21, "transformPlan"
'4.Generowanie imiennych kart pracy
karty_pracy_generator.gen



Application.DisplayAlerts = False
ActiveWorkbook.Save
Application.DisplayAlerts = True
main.progress_bar 100, "FINITO"
Unload gen_load_form
MsgBox "Karty pracy gotowe. Uzupe³nij specjalne "
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Workbooks("Karty Pracy - " + source_name + ".xlsx").Activate
ActiveWorkbook.Worksheets(1).Select
'GeneratorPlanuPracy.generator.Cells(55, 11) = Timer - Start
End Sub

'Public Sub generuj_time()
''Prod source_path
'Start = Timer
'source_path = ActiveWorkbook.Path
'
''1 i 2 .Skopiowanie Ÿród³a do nowego arkusza i go Utworzenie  - Karty pracy - Nazwa Ÿród³a
'create_new_sheet "Karty Pracy - " + source_name
'sourceWS.Range("K32") = Timer - Start
''3.Transformacja arksza pod karty pracy
'
'transformPlan
'
''4.Generowanie imiennych kart pracy - ju¿ zrobione tylko zaimplementowaæ
'karty_pracy_generator.gen
'
'
'Unload karty_pracy_dlgbox
'Workbooks("Karty Pracy - " + source_name).Activate
'ActiveWorkbook.Worksheets(1).Select
'Application.DisplayAlerts = False
'ActiveWorkbook.Save
'Application.DisplayAlerts = True
'MsgBox "Karty pracy gotowe. Uzupe³nij specjalne "
'End Sub

Public Sub show_dlgbox()
karty_pracy_dlgbox.Show
End Sub

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
