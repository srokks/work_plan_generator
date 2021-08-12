Attribute VB_Name = "main"
Private sourceWS As Worksheet
Private workws As Worksheet
Private sheet_name As String
Private cur_month, cur_year As String
Private weekend_color, working_hours As Long
Private height, width As Integer
Private sheet_name_state As Boolean

Public Sub workshitapostasis()
Dim TempSheetName As String
TempSheetName = UCase("generator")
For Each Sheet In Worksheets
    If UCase("generator") = UCase(Sheet.Name) Or UCase("test") = UCase(Sheet.Name) Then
    Else
        Application.DisplayAlerts = False
        Worksheets(Sheet.Name).Delete
        Application.DisplayAlerts = True
    End If
Next
MsgBox "AREA CLEARED", vbOKOnly, "TO YOUR COMMAND"
End Sub


Sub unlock_source(ByVal strPassword As String, Optional ByVal v As Boolean)
If v Then
    sourceWS.Unprotect strPassword
    Exit Sub
End If
If sourceWS.ProtectContents Then
    sourceWS.Unprotect strPassword
Else
    sourceWS.Protect strPassword, DrawingObjects:=False, Contents:=True, Scenarios:= _
            False
End If
End Sub

Sub set_color(ByVal kolor As Long)
weekend_color = kolor
update_form
update_sheet
End Sub

Sub set_weekends()
Application.Calculation = xlCalculationAutomatic
For i = 3 To width Step 2
    temp_text = workws.Range(Cells(2, i), Cells(2, i)).Text
    If temp_text = "nd" Or temp_text = "sb" Then
        workws.Range(Cells(2, i), Cells(height, i + 1)).Interior.Color = weekend_color
    End If
    workws.Range(Cells(2, i), Cells(2, i)).Select
Next i
Application.Calculation = xlCalculationManual
End Sub

Sub generuj()
Unload gen_plan_dlgbox
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
gen_load_form.progres_label.width = 0
gen_load_form.Show
Start = Timer
'test_data
progress_bar 0.28, "test_data"
create_sheet
progress_bar 0.47, "create_sheet"
fill_empl
progress_bar 7.68, "fill_empl"
fill_days
progress_bar 18.94, "fill_days"
fill_work_days
progress_bar 46.46, "fill_work_days"
set_night_cond_form
progress_bar 71.26, "set_night_cond_form"
set_weekends
progress_bar 76.04, "set_weekend_cond_form"
fill_work_hours
progress_bar 87.08, "fill_work_hours"
fill_headers
progress_bar 88.8
fill_downinfo
progress_bar 89.6
set_print_info
progress_bar 91.2
'set_workareas
progress_bar 100
Unload gen_load_form
MsgBox "GOTOWE"
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub


Sub set_workareas()
width = 66
height = 39
For i = 3 To width Step 2
    If Range(Cells(2, i), Cells(2, i)) = "nd" Or Range(Cells(2, i), Cells(2, i)) = "sb" Then
        Range(Cells(2, i), Cells(height, i)).Select
    End If
    
Next
End Sub

Sub fill_work_hours()
Dim temp_text As String

For i = 0 To 1
    width = width + 1
    For j = 4 To height Step 2
        If i = 0 Then
            temp_text = "=SUMA(" + workws.Range(Cells(j, 3), Cells(j, width - 1)).Address(RowAbsolute = True, ColumnAbsolute = False) + ")"
            workws.Range(Cells(j, width), Cells(j + 1, width)).Select
            With Selection
                .Merge
                .Borders.LineStyle = xlContinuous
                .Borders.Color = vbBlack
                .Borders.Weight = xlThin
                .font.Name = "Cambria"
                .font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .font.Color = -16776961
                .ColumnWidth = 8.11
                .FormulaLocal = temp_text
                .NumberFormat = "0"
            End With
        End If
        If i = 1 Then
            temp_text = "=" + workws.Range(Cells(j, width - 1), Cells(j, width - 1)).Address(RowAbsolute = True, ColumnAbsolute = False) + "-" + workws.Range("A2").Address
            workws.Range(Cells(j, width), Cells(j + 1, width)).Select
            With Selection
                .Merge
                .Borders.LineStyle = xlContinuous
                .Borders.Color = vbBlack
                .Borders.Weight = xlThin
                .font.Name = "Cambria"
                .font.Size = 8
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .font.Color = -16776961
                .ColumnWidth = 8.11
                .FormulaLocal = temp_text
                .NumberFormat = "0"
            End With
        End If
        
    Next j
'   width = width + 1
Next i

End Sub

Sub fill_headers()
workws.Range("A1") = "Dzia³ " + sourceWS.Range("c2")
If (working_hours = 0) Or (cur_year <> Year(Now)) Then
    workws.Range("A2").Select
    With Selection
        .AddComment
        .Comment.Visible = False
        .Comment.Text Text:= _
            "***GENERATOR***:" & Chr(10) & "UZUPE£NIJ RÊCZNIE ILOŒÆ GODZIN!"
        .Comment.Visible = True
    workws.Range("A2") = working_hours
    workws.Range("A2").NumberFormat = "#,##0_ ;-#,##0 "
    End With
End If
workws.Range("A2") = working_hours
workws.Range("A2").NumberFormat = "0"
workws.Range("A3") = "LP"
workws.Range("B2") = cur_month
workws.Range(Cells(1, width - 1), Cells(1, width - 1)) = "iloœæ godzin w miesi¹cu"
workws.Range(Cells(1, width), Cells(1, width)) = "iloœæ nadgodzin w miesi¹cu"
workws.Range(Cells(1, 1), Cells(1, width - 2)).Select
With Selection
    .Merge
    .Borders.LineStyle = xlContinuous
    .Borders.Color = vbBlack
    .Borders.Weight = xlThin
    .font.Name = "Cambria"
    .font.Size = 10
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = 14277081
End With
workws.Range(Cells(2, 1), Cells(2, 2)).Select
With Selection
    .Borders.LineStyle = xlContinuous
    .Borders.Color = vbBlack
    .Borders.Weight = xlThin
    .font.Name = "Cambria"
    .font.Size = 10
    .font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = 14277081
End With
workws.Range(Cells(3, 1), Cells(3, 2)).Select
With Selection
    .Borders.LineStyle = xlContinuous
    .Borders.Color = vbBlack
    .Borders.Weight = xlThin
    .font.Name = "Cambria"
    .font.Size = 7
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = 14277081
End With
workws.Range(Cells(1, width - 1), Cells(3, width - 1)).Select
With Selection
    .Merge
    .Borders.LineStyle = xlContinuous
    .Borders.Color = vbBlack
    .Borders.Weight = xlThin
    .font.Name = "Cambria"
    .font.Size = 8
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = 14277081
    .WrapText = True
End With
workws.Range(Cells(1, width), Cells(3, width)).Select
With Selection
    .Merge
    .Borders.LineStyle = xlContinuous
    .Borders.Color = vbBlack
    .Borders.Weight = xlThin
    .font.Name = "Cambria"
    .font.Size = 8
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = 14277081
    .WrapText = True
End With
End Sub

Sub create_sheet()
If WorksheetExists(sheet_name) Then
    Application.DisplayAlerts = False
    Worksheets(sheet_name).Delete
    Application.DisplayAlerts = True
End If
Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = sheet_name
Set workws = Worksheets(sheet_name)


End Sub

Sub set_print_info()
    Application.PrintCommunication = True
    ActiveWindow.Zoom = 78
    With workws.PageSetup
        .Zoom = False
        .PrintArea = workws.Range(Cells(1, 1), Cells(height, width)).Address
        .Orientation = xlLandscape
        .FitToPagesTall = 1
        .FitToPagesWide = 1
        .LeftHeader = "&F"
        .RightHeader = "wygenerowano: &D;&T"
    End With
    
If printer_a3_comp Then
    workws.PageSetup.PaperSize = xlPaperA3
Else
    workws.PageSetup.PaperSize = xlPaperA4
End If
Application.PrintCommunication = False
Range("C4:D4").Select
'ActiveWindow.FreezePanes = True
'workWS.Range("A1").Select
End Sub

Sub progress_bar(ByVal i As Long, Optional ByVal s As String, Optional ByVal grow As Boolean)
If grow = True Then
    gen_load_form.progres_label.width = gen_load_form.progres_label.width + i
ElseIf grow = False Then
    gen_load_form.progres_label.width = gen_load_form.progres_label.width + i * (gen_load_form.frame_progress.width * 0.01)
End If
If s <> "" Then
    gen_load_form.progres_label.Caption = s
End If
gen_load_form.Repaint
End Sub



Sub fill_downinfo()
height = height + 3
workws.Range(Cells(height, 1), Cells(height, width)).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .AddIndent = False
        .ShrinkToFit = True
        .MergeCells = True
        .font.Bold = False
        .font.Size = 16
        .font.Name = "Cambria"
        .Interior.Color = 6299648
        .font.Color = vbWhite
        .WrapText = True
        .RowHeight = 50.03
        .Value = "Harmonogram pracy w uzasadnionych przypadkach mo¿e ulec zmianie. O ka¿dej zmianie Pracownik zostanie poinformowany na min.7 dni przed dniem ,w którym nast¹pi zmiana lub w sytuacjach skrajnych niezale¿nych od Pracodawcy-najpóŸniej do dniówki roboczej poprzedzaj¹cej dzieñ, w którym nast¹pi zmiana"

End With

End Sub

Sub fill_work_days()
Dim temp_range As Range
Dim i, j As Integer
Dim temp_text As String
width = get_width(cur_month)
'pêtla wype³nia pole godzin odpowiedni¹ formu³om dla ka¿dego pracownika
For j = 4 To height Step 2
    For i = 3 To width Step 2
        Set temp_range = workws.Range(Cells(j, i), Cells(j, i + 1))
        adr = temp_range.Offset(1, 0).Address(RowAbsolute = True, ColumnAbsolute = True)
        temp_text = "=JE¯ELI(LUB(" + left(adr, (Len(adr) - 1) / 2) + "=0;" + Right(adr, (Len(adr) - 1) / 2) + "=0);0;JE¯ELI(" + Right(adr, (Len(adr) - 1) / 2) + ">" + left(adr, (Len(adr) - 1) / 2) + ";" + Right(adr, (Len(adr) - 1) / 2) + "-" + left(adr, (Len(adr) - 1) / 2) + ";JE¯ELI(" + Right(adr, (Len(adr) - 1) / 2) + "<" + left(adr, (Len(adr) - 1) / 2) + ";24-" + left(adr, (Len(adr) - 1) / 2) + "+" + Right(adr, (Len(adr) - 1) / 2) + ";""B£AD"")))"
        With temp_range
            .Merge
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
            .FormulaLocal = temp_text
            .font.Bold = False
            .font.Size = 8
            .font.Name = "Cambria"
            .ColumnWidth = 2.56
        End With
    Next i
Next j
'wype³nienie obramowania komórek dni
'-------------------------
workws.Range(Cells(4, 3), Cells(height, width)).Select
With Selection
        .Borders.LineStyle = xlContinuous
        .Borders.Color = vbBlack
        .Borders.Weight = xlThin
        .font.Name = "Cambria"
        .font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
End With
End Sub


Sub fill_days()

Dim temp_range As Range
Dim i, j As Integer
Dim temp_text As String
Dim width As Integer
Dim temp_date, date_separator As String
width = get_width(cur_month)
date_separator = Application.International(xlDateSeparator)

month_date = Month("01/" + cur_month + "/" + cur_year)
Select Case Application.International(xlDateOrder)
    Case 0
        'MsgBox "Month -Day - Year"
        temp_date = CStr(month_date) + date_separator + "1" + date_separator + cur_year
    Case 1
        'MsgBox "Day - Month - Year"
        temp_date = "1" + date_separator + CStr(month_date) + date_separator + cur_year
    Case 2
        'MsgBox "Year - Month - Day"
        temp_date = cur_year + date_separator + CStr(month_date) + date_separator + "1"
End Select
For i = 3 To width Step 2
    For j = 2 To 3
        workws.Range(Cells(j, i), Cells(j, i + 1)).Select
        With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = True
                .font.Bold = True
                .font.Size = 10
                .font.Name = "Cambria"
                .Borders.LineStyle = xlContinuous
                .Borders.Color = vbBlack
                .Borders.Weight = xlThin
                .ColumnWidth = 2.56
            End With
            If j = 2 Then
                Selection.FormulaR1C1 = _
            "=CHOOSE(WEEKDAY(R[1]C,2),""pn"",""wt"",""œr"",""czw"",""pt"",""sb"",""nd"")"

            ElseIf j = 3 And i = 3 Then
                Selection.FormulaLocal = temp_date
                Selection.NumberFormat = "d"
            ElseIf j = 3 And i <> 3 Then
                Selection.FormulaR1C1 = "=RC[-2]+1"
            End If
            
    Next j
Next i
End Sub




Sub fill_empl()
Dim empl_list() As String
empl_list = gen_emp
j = 4
temp_int = 1
'wpisanie do arkusza pracowników+wype³nienie komórek + merge
For Each row In empl_list
    If Not row = "" Then
            workws.Cells(j, 2).Select
            With Selection
                .font.Name = "Cambria"
                .NumberFormat = "@"
                .font.Bold = False
                .font.Size = 11
                .font.Name = "Cambria"
                .Interior.Color = 14277081
            End With
        If flag Then
            workws.Cells(j, 2) = row
            flag = False
            workws.Range(Cells(j - 1, 1), Cells(j, 1)).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .MergeCells = True
                .Value = CStr(temp_int) + "."
                .NumberFormat = "@"
                .font.Bold = False
                .font.Size = 7
                .font.Name = "Cambria"
                .Interior.Color = 14277081
                .Borders.LineStyle = xlContinuous
                .Borders.Color = vbBlack
                .Borders.Weight = xlThin
             End With
             If j Mod 2 Then
                temp_int = temp_int + 1
            End If
        Else
            workws.Cells(j, 2) = row
            flag = True
        End If
        j = j + 1
    End If
Next row
height = get_height(empl_list)
'MsgBox down_zakres
For i = 4 To height - 1 Step 2
    workws.Range(Cells(i, 2), Cells(i + 1, 2)).Select
    With Selection
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Next i
workws.Columns(2).ColumnWidth = 26.33
End Sub


Private Sub test_data()
cur_month = "Styczeñ"
cur_year = "2021"
sheet_name = "testy"
weekend_color = 11851260
If WorksheetExists(sheet_name) Then
    Application.DisplayAlerts = False
    Worksheets(sheet_name).Delete
    Application.DisplayAlerts = True
End If
Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = sheet_name
Set workws = Worksheets(sheet_name)
End Sub

Function get_height(empl_list) As Integer
get_height = GetArrLength(empl_list) * 2
get_height = get_height + 1
End Function

Function gen_emp() As String()
Set sourceWS = Worksheets("Generator")
Dim empl_list() As String
ReDim empl_list(1 To 2, 1 To 1)
Dim i As Integer: i = 4
Do While ((Not IsEmpty(sourceWS.Cells(i, 3)) Or (Not IsEmpty(sourceWS.Cells(i + 1, 3)))))
    
    If i = 4 Then
        empl_list(1, i - 3) = sourceWS.Cells(i, 3)
        empl_list(2, i - 3) = sourceWS.Cells(i + 1, 3)
        ReDim Preserve empl_list(1 To 2, 1 To i - 2)
    Else
        empl_list(1, i / 2 - 1) = sourceWS.Cells(i, 3)
        empl_list(2, i / 2 - 1) = sourceWS.Cells(i + 1, 3)
        ReDim Preserve empl_list(1 To 2, 1 To i / 2)
    End If
   i = i + 2
Loop

'MsgBox GetArrLength(empl_list)
'For i = 1 To GetArrLength(empl_list) 'wyœwietlanie listy pracowników
'MsgBox (CStr(i) + "." + empl_list(1, i) + " " + empl_list(2, i))
'Next i
gen_emp = empl_list
End Function
Private Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrSize_2D = 0
   Else
      X = UBound(a, 1) - LBound(a, 1) + 1
      Y = UBound(a, 2) - LBound(a, 2) + 1
      GetArrLength = Y
   End If
End Function
Function get_width(ByVal cur_month As String) As Integer
Dim month_date As String: month_date = Month("01-" + cur_month + "-2021")
If nr_miesiaca Then
    get_width = mon_lp
Else

Select Case month_date
        Case 1, 3, 5, 7, 8, 10, 12
            get_width = 31
        Case 2
            If ((CInt(Year(month_date)) Mod 4) = 0) Then
              get_width = 28
            Else
                get_width = 29
            End If
        Case 4, 6, 9, 11
            get_width = 30
        Case Else
            MsgBox "B³êdny miesi¹c", vbCritical
    End Select
End If
get_width = get_width * 2 + 2
End Function

Function WorksheetExists(SheetName As String) As Boolean
Dim TempSheetName As String
TempSheetName = UCase(SheetName)
WorksheetExists = False
For Each Sheet In Worksheets
    If TempSheetName = UCase(Sheet.Name) Then
        WorksheetExists = True
        Exit Function
    End If
Next Sheet
End Function


Sub init_form()
Set sourceWS = Worksheets("Generator")
'unlock_source "chelsea77" 'unlock
weekend_color = 11851260
sheet_name_state = False
gen_plan_dlgbox.choose_year_cmbbox.AddItem Year(Now())
gen_plan_dlgbox.choose_year_cmbbox.AddItem Year(Now()) + 1
gen_plan_dlgbox.choose_year_cmbbox.ListIndex = 0
End Sub

Sub update_sheet()
'unlock_source "chelsea77", True
sourceWS.Range("K4") = cur_month
cur_month = gen_plan_dlgbox.choose_month_cmbbox.Value


cur_year = gen_plan_dlgbox.choose_year_cmbbox.Value
sourceWS.Range("K5") = cur_year

sourceWS.Range("K6") = weekend_color
sourceWS.Range("K6").Interior.Color = weekend_color
weekend_color = sourceWS.Range("K6")

If sheet_name_state = False Then
    sheet_name = cur_month + " " + cur_year
    
    If WorksheetExists(sheet_name) Or count_worksheet_copies(sheet_name) > 1 Then
        sheet_name = cur_month + " " + cur_year + " (" + CStr(count_worksheet_copies(sheet_name)) + ")"
    End If
Else
    sheet_name_state = False
End If



'If check_count(sheet_name) > 1 Then
'     For i = 0 To check_count(she)
'
'    sheet_name = sheet_name + " (" + CStr(check_count(sheet_name + " (" + CStr(check_count(sheet_name) + 1) + ")")) + ")"
'
'    MsgBox ("Nazwa skoroszytu zdublikowana" & vbNewLine & "Zostanie zapisane pod nazw¹:" & vbNewLine & sheet_name)
'End If
sourceWS.Range("K7") = sheet_name
If cur_month <> "" Then
    working_hours = sourceWS.Range("F4:f15").Find(cur_month).Offset(, 1)
End If
sourceWS.Range("K8") = working_hours

update_form
'unlock_source "chelsea77"
'TODO:odkomentuj blokadê w trzech miejscach wyszukaj u¿ycie proc unlock_source
End Sub

Sub update_form()
gen_plan_dlgbox.choose_month_cmbbox.Value = cur_month
gen_plan_dlgbox.choose_color_btn.BackColor = weekend_color
gen_plan_dlgbox.sheet_name_txtbox.Value = sheet_name
If (gen_plan_dlgbox.choose_month_cmbbox.Value = "") Or (gen_plan_dlgbox.choose_year_cmbbox.Value = "") Or (gen_plan_dlgbox.choose_color_btn.BackColor = 0) Then
    Else
        gen_plan_dlgbox.gen_btn.Enabled = True
    End If
    
End Sub

Sub source_btn_click()
Load gen_load_form
gen_plan_dlgbox.Show
End Sub

Sub update_sheet_name()
sheet_name = gen_plan_dlgbox.sheet_name_txtbox.Text
sheet_name_state = True
update_sheet
End Sub
Sub prev_data()
Dim sheet_chng_state As Boolean: sheet_chng_state = False
Dim sheet_name_chgstate As Boolean
cur_month = sourceWS.Range("K4")
cur_year = Year(Now())
weekend_color = sourceWS.Range("K6")
working_hours = sourceWS.Range("K12")
sheet_name = cur_month + " " + cur_year

update_form
End Sub

Function count_worksheet_copies(ByVal sheet_name As String) As Integer
Dim i As Integer: i = 1
Dim TempSheetName As String
TempSheetName = UCase(sheet_name)
For Each Sheet In Worksheets
    If TempSheetName = UCase(Sheet.Name) Then
        TempSheetName = UCase(sheet_name + " (" + CStr(i + 1) + ")")
        i = i + 1
    End If
Next Sheet
count_worksheet_copies = i

End Function


Sub set_night_cond_form()
Dim workws As Worksheet
Set workws = Worksheets(sheet_name)
Dim workRange As Range
Dim adres_a, adres_b As String

Dim col, row As Long
For row = 4 To height Step 2
    For col = 3 To width Step 2
        adres_a = workws.Cells(row, col).Offset(1, 0).Address
        adres_b = workws.Cells(row, col - 1).Offset(1, 2).Address
        temp_adres = "=JE¯ELI(LUB(" + adres_a + "=0;" + adres_b + "=0);FA£SZ;JE¯ELI(" + adres_a + ">" + adres_b + ";PRAWDA;JE¯ELI(LUB(0>" + adres_a + ";" + adres_a + "<7);PRAWDA;FA£SZ)))"
        workws.Range(Cells(row, col), Cells(row, col).Offset(1, 0)).Select
            Selection.FormatConditions.Add Type:=xlExpression, Formula1:=temp_adres
            Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).font
            .Color = -16776961
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    Next col
Next row
End Sub

'Function WorksheetExists(SheetName As String) As Boolean
'Dim TempSheetName As String
'TempSheetName = UCase(SheetName)
'WorksheetExists = False
'For Each Sheet In Worksheets
'    If TempSheetName = UCase(Sheet.Name) Then
'        WorksheetExists = True
'        Exit Function
'    End If
'Next Sheet
'End Function
'

Sub test_time()
test_data
'sourceWS.Range("K19") = Timer - Start
create_sheet
'sourceWS.Range("K20") = Timer - Start
fill_empl
'sourceWS.Range("K21") = Timer - Start
fill_days
'sourceWS.Range("K22") = Timer - Start
fill_work_days
'sourceWS.Range("K23") = Timer - Start
set_night_cond_form
'sourceWS.Range("K24") = Timer - Start
set_weekend_cond_form
'sourceWS.Range("K25") = Timer - Start
fill_work_hours
'sourceWS.Range("K26") = Timer - Start
fill_headers
'sourceWS.Range("K27") = Timer - Start
fill_downinfo
'sourceWS.Range("K28") = Timer - Start
set_print_info
'sourceWS.Range("K29") = Timer - Start
End Sub


