Attribute VB_Name = "karty_pracy_generator"
Public sheet_count As Integer
Private prev_sheetName, comp_name As String
Private height As Integer
Private Sub gen_workCard(ByVal empl As String)
Worksheets.Add(After:=Sheets(Sheets.Count)).Name = CStr(sheet_count) + "." + empl
prev_sheetName = CStr(sheet_count) + "." + empl
fillWorkCard prev_sheetName
End Sub

Private Sub set_print_info(ByVal workWS_name As String)
Dim workws As Worksheet
Set workws = Worksheets(workWS_name)
Application.PrintCommunication = True
ActiveWindow.Zoom = 75
With workws.PageSetup
    .Zoom = False
    .PrintArea = workws.Range(Cells(1, 1), Cells(height, 12)).Address
    .Orientation = xlPortrait
    .FitToPagesTall = 1
    .FitToPagesWide = 1
    .PaperSize = xlPaperA4
End With
Application.PrintCommunication = False
End Sub
    

Private Sub fillDownInfo(ByVal workWS_name As String)
Dim workws As Worksheet
Set workws = Worksheets(workWS_name)
comp_name = karty_pracy.comp_name
height = height + 2
workws.Range(workws.Cells(height, 1), workws.Cells(height, 3)).Select
With Selection
        .Merge
        .font.Name = "Calibri"
        .font.Size = 10
        .font.Bold = False
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    .Value = "podpis kierownika jednostki:"
End With
height = height + 3
workws.Range(workws.Cells(height, 1), workws.Cells(height + 1, 12)).Select
With Selection
        .Merge
        .font.Name = "Times New Roman"
        .font.Size = 7
        .font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Value = "Oznaczenia: wp.-wolne za nadgodziny, Uw.-urlop wypoczynkowy, Uok.-urlop okolicznoœciowy, nn-nieob.nieuspraw., nu.-nieobecnoœæ usprawiedliwiona,U.wych.-urlop wychowawczy, Uop-opieka naz zdrowym dzieckiem do lat 14"
End With
height = height + 3
workws.Range(workws.Cells(height, 7), workws.Cells(height, 9)).Select
With Selection
        .Merge
        .font.Name = "Calibri"
        .font.Size = 10
        .font.Bold = False
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    .Value = "podpis osoby zarz¹dzaj¹cej:"
End With
height = height + 10
workws.Range(workws.Cells(height, 1), workws.Cells(height, 12)).Select
With Selection
        .Merge
        .font.Name = "Times New Roman"
        .font.Size = 10
        .font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = comp_name
End With
End Sub

Private Sub fillWorkCard(ByVal workWS_name As String)
Dim workws As Worksheet
Set workws = Worksheets(workWS_name)
set_columnWidth workWS_name
fillHeadInfo workWS_name
fillHours workWS_name
fillDownInfo workWS_name
set_print_info workWS_name
workws.Range("a1").Select
End Sub

Private Sub fillHours(ByVal workWS_name As String)
Dim workws As Worksheet
Dim empl_modif As Integer
Dim temp_adr, temp_formula As String
Set workws = ActiveWorkbook.Worksheets(workWS_name)
Set daneWS = ActiveWorkbook.Worksheets("Harmonogram pracy")
Dim i, k As Integer: i = 3
k = 14
Do While daneWS.Range(daneWS.Cells(4, i), daneWS.Cells(4, i)) <> ""
    'daneWS.Range(Cells(4, i), Cells(4, i)).Select
    For j = 1 To 14
        empl_modif = ((sheet_count - 1) * 2)
        Select Case j
            Case Is = 1
                'nr dnia
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)) = daneWS.Range(daneWS.Cells(4, i), daneWS.Cells(4, i))
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)).NumberFormat = "d"
            Case Is = 2
                'rozpoczêcie pracy
                
                temp_adr = "'Harmonogram pracy'!" + daneWS.Range(daneWS.Cells(6 + empl_modif, i), daneWS.Cells(6 + empl_modif, i)).Address
                temp_formula = "=Je¿eli(Lub(" + temp_adr + "=""w5"";" + temp_adr + "=""ws"";" + temp_adr + "=""wn"";" + temp_adr + "="""");"""";" + temp_adr + ")"
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)).FormulaLocal = temp_formula
                
            Case Is = 3
                'zakoñczenie pracy
                temp_adr = "'Harmonogram pracy'!" + daneWS.Range(daneWS.Cells(6 + empl_modif, i + 1), daneWS.Cells(6 + empl_modif, i + 1)).Address
                temp_formula = "=Je¿eli(Lub(" + temp_adr + "=""w5"";" + temp_adr + "=""ws"";" + temp_adr + "=""wn"";" + temp_adr + "="""");"""";" + temp_adr + ")"
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)).FormulaLocal = temp_formula
            Case Is = 4
                '³¹czny czas pracy
                temp_adr = "'Harmonogram pracy'!" + daneWS.Range(daneWS.Cells(5 + empl_modif, i), daneWS.Cells(5 + empl_modif, i + 1)).Address
                temp_formula = "=JE¯ELI(SUMA(" + temp_adr + ")=0;"""";SUMA(" + temp_adr + "))"
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)).FormulaLocal = temp_formula
            Case Is = 7
                'godziny nocne
                temp_adr = workws.Range(workws.Cells(k, 2), workws.Cells(k, 2)).Address
                temp_adr2 = workws.Range(workws.Cells(k, 3), workws.Cells(k, 3)).Address
                temp_formula = "=JE¯ELI(ORAZ(CZY.LICZBA(" + temp_adr + ");CZY.LICZBA(" + temp_adr2 + "));JE¯ELI(" + temp_adr + ">" + temp_adr2 + " ;(24- " + temp_adr + ")-JE¯ELI(22-" + temp_adr + "<=0;0;22-" + temp_adr + " )+JE¯ELI(" + temp_adr2 + " >6;6;" + temp_adr2 + " );JE¯ELI(ORAZ(" + temp_adr2 + " >22;" + temp_adr2 + "<=24);" + temp_adr2 + " -22;"" ""));"" "")"
                workws.Range(workws.Cells(k, j), workws.Cells(k, j)).FormulaLocal = temp_formula
        End Select
    Next j
     k = k + 1
i = i + 2

Loop
workws.Range(Cells(14, 1), Cells(k - 1, 1)).Select
With Selection
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = -4.99893185216834E-02
        .Interior.PatternTintAndShade = 0
        .Borders.LineStyle = xlContinuous
        .Borders.ColorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .font.Name = "Calibri"
        .font.Size = 8
        .font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
workws.Range(Cells(14, 2), Cells(k - 1, 12)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
        .Borders.ColorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .font.Name = "Calibri"
        .font.Size = 8
        .font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

'dolne sumy godzin i obramowanie ich
workws.Range(Cells(k, 1), Cells(k, 12)).Select
With Selection.Borders
    .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
End With
temp_formula = "=SUM(R[-" + CStr(k - 14) + "]C:R[-1]C)"
workws.Range(Cells(k, 4), Cells(k, 4)).FormulaR1C1 = temp_formula
workws.Range(Cells(k, 4), Cells(k, 4)).AutoFill Destination:=Range(Cells(k, 4), Cells(k, 8)), Type:=xlFillDefault
height = k


End Sub

Private Sub set_columnWidth(ByVal workWS_name As String)
    Dim workws As Worksheet
    Set workws = Worksheets(workWS_name)
    workws.Range("A1").ColumnWidth = 4.89
    workws.Range("B1").ColumnWidth = 9.11
    workws.Range("C1").ColumnWidth = 9.33
    workws.Range("D1").ColumnWidth = 9.56
    workws.Range("E1").ColumnWidth = 6.33
    workws.Range("F1").ColumnWidth = 13.11
    workws.Range("G1").ColumnWidth = 6.89
    workws.Range("H1").ColumnWidth = 8.78
    workws.Range("I1").ColumnWidth = 8.22
    workws.Range("J1").ColumnWidth = 6.89
    workws.Range("K1").ColumnWidth = 5.78
    workws.Range("L1").ColumnWidth = 4.89
End Sub


Public Sub fillHeadInfo(ByVal workWS_name As String)

Dim workws, daneWS As Worksheet
Set workws = Worksheets(workWS_name)


workws.Range("A1").Select
With Selection
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = True
    .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    .Value = "KARTA EWIDENCJI  CZASY PRACY ZA MIESI¥C:"
End With
workws.Range("E1").Select
With Selection
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Value = karty_pracy.cur_month
End With
workws.Range("i1").Select
With Selection
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = True
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Value = "ROK"
End With
workws.Range("j1").Select
With Selection
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = True
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .Value = karty_pracy.cur_year
End With
workws.Range("H3:I3").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .Value = "Norma miesiêczna:"
End With
workws.Range("J3").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .Formula = karty_pracy.hours
End With
workws.Range("A3:C3").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .Value = "Imiê i nazwisko:"
End With

Set daneWS = Worksheets("HARMONOGRAM PRACY")

empl_modif = ((sheet_count - 1) * 2)
temp_formula = "='Harmonogram pracy'!" + daneWS.Range(daneWS.Cells(5 + empl_modif, 2), daneWS.Cells(5 + empl_modif, 2)).Address
workws.Range("D3").FormulaLocal = temp_formula
workws.Range("D3:E3").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
temp_formula = "='Harmonogram pracy'!" + daneWS.Range(daneWS.Cells(6 + empl_modif, 2), daneWS.Cells(6 + empl_modif, 2)).Address
workws.Range("D4").FormulaLocal = temp_formula
workws.Range("D4:E4").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
workws.Range("A4:C4").Select
With Selection
    .Merge
    .font.Name = "Calibri"
    .font.Size = 9
    .font.Bold = False
    .HorizontalAlignment = xlRight
    .VerticalAlignment = xlCenter
    .Value = "System czasu pracy:"
End With
workws.Range("A1:D1").Merge
workws.Range("E1:F1").Merge
fillHeaders workWS_name
set_columnWidth workWS_name

End Sub

Private Sub fillHeaders(ByVal workWS_name As String)

Dim workws As Worksheet
Set workws = Worksheets(workWS_name)
workws.Range("A6:A13").Select
With Selection
    .Merge
    .Value = "Dzieñ m-ca"
End With
workws.Range("B6:C12").Select
With Selection
    .Merge
    .Value = "Rzeczywisty czas pracy"
End With
workws.Range("B13").Merge
workws.Range("B13").Value = "Rozpoczêcie pracy"
workws.Range("c13").Merge
workws.Range("c13").Value = "Zakoñczenie pracy"
workws.Range("D6:d13").Merge
workws.Range("D6:d13").Value = "£¹czny czas pracy lub symbol nieobecnoœci"
workws.Range("E6:E13").Merge
workws.Range("E6:E13").Value = "Godziny urlopu"
workws.Range("F6:F13").Merge
workws.Range("F6:F13").Value = "Zwolnienia od pracy oraz inne uspr. I nieuspr. nieobecnoœci"
workws.Range("G6:G13").Merge
workws.Range("G6:G13").Value = "Godziny nocne"
workws.Range("H6:k12").Merge
workws.Range("h6:k12").Value = "Godziny pracy dodatkowej"
workws.Range("h13").Value = "Godziny nadliczbowe"
workws.Range("i13").Value = "W niedziele i œwiêta"
workws.Range("j13").Value = "W dniu wolnym"
workws.Range("k13").Value = "Dy¿ury"
workws.Range("L6:L13").Merge
workws.Range("L6:L13").Value = "Uwagi"

workws.Range("A6:L13").Select
    With Selection
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = -4.99893185216834E-02
        .Interior.PatternTintAndShade = 0
        .Borders.LineStyle = xlContinuous
        .Borders.ColorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
        .font.Name = "Calibri"
        .font.Size = 9
        .font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub



Public Sub gen()
sheet_count = 1
Dim sourceWS As Worksheet
Set sourceWS = Worksheets("HARMONOGRAM PRACY")
'Do While ((Not IsEmpty(sourceWS.Cells(i, 3)) Or (Not IsEmpty(sourceWS.Cells(i + 1, 3)))))
Dim i As Integer
i = 5
Do While ((Not IsEmpty(sourceWS.Cells(i, 2)) Or (Not IsEmpty(sourceWS.Cells(i + 1, 3)))))
    gen_workCard (sourceWS.Cells(i, 2).Value + " " + sourceWS.Cells(i + 1, 2).Value)
    'GeneratorPlanuPracy.generator.Cells(34 + sheet_count, 11) = Timer - Start
    main.progress_bar 1.5, sourceWS.Cells(i, 2).Value + " " + sourceWS.Cells(i + 1, 2), True
    sheet_count = sheet_count + 1
    i = i + 2
    If sourceWS.Cells(i, 2).Value = "WS - dzieñ wolny za œwiêto" Then
    Exit Sub
    End If
Loop


End Sub
Private Sub usun()

For Each Sheet In Worksheets
    If UCase(Sheet.Name) = UCase("Harmonogram Pracy") Or UCase(Sheet.Name) = UCase("1") Or UCase(Sheet.Name) = UCase("DANE WEJŒCIOWE") Or UCase(Sheet.Name) = UCase("Harmonogram Pracy wzór") Then
    
    Else
        Application.DisplayAlerts = False
        Worksheets(Sheet.Name).Delete
        Application.DisplayAlerts = True
    End If
Next
End Sub
