' ============================================================
' MODULO VBA — Dashboard Meteo Stazioni
' Incolla in un modulo VBA del file .xlsm
' Strumenti > Macro > Visual Basic Editor > Insert > Module
' ============================================================

Option Explicit

' ---- COSTANTI ----
Private Const WIND_COL  As String = "TOP 92;wind_speed;Avg (m/s)"
Private Const TEMP_COL  As String = "TEMP-UMID;temperature;Avg (°C)"
Private Const HUM_COL   As String = "TEMP-UMID;humidity;Avg (%)"
Private Const PRES_COL  As String = "GEOVES BOX;air_pressure;Avg (hPa)"
Private Const W88_COL   As String = "RIF 88;wind_speed;Avg (m/s)"
Private Const W70_COL   As String = "RIF 70;wind_speed;Avg (m/s)"
Private Const W50_COL   As String = "RIF 50;wind_speed;Avg (m/s)"

' ============================================================
' ENTRY POINT — eseguire questa macro
' ============================================================
Sub ImportaCSVECreaaDashboard()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim cartella As String
    cartella = SelezionaCartella()
    If cartella = "" Then Exit Sub

    ' 1. Importa tutti i CSV dalla cartella
    Call ImportaCSV(cartella)

    ' 2. Calcola statistiche mensili
    Call CalcolaStatistiche()

    ' 3. Crea dashboard visiva
    Call CreaDashboard()

    ' 4. Esporta JSON per Streamlit
    Call EsportaJSON()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "✅ Import completato!" & vbCrLf & _
           "- Foglio 'Dati Raw': tutti i CSV uniti" & vbCrLf & _
           "- Foglio 'Statistiche': medie mensili" & vbCrLf & _
           "- Foglio 'Dashboard': grafici" & vbCrLf & _
           "- File 'export_streamlit.json' creato", _
           vbInformation, "Dashboard Meteo"
End Sub

' ============================================================
' SELEZIONE CARTELLA
' ============================================================
Function SelezionaCartella() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Seleziona la cartella con i file CSV"
    fd.InitialFileName = ThisWorkbook.Path
    If fd.Show = -1 Then
        SelezionaCartella = fd.SelectedItems(1) & "\"
    Else
        SelezionaCartella = ""
    End If
End Function

' ============================================================
' IMPORT CSV
' ============================================================
Sub ImportaCSV(cartella As String)
    ' Crea o svuota foglio Dati Raw
    Dim wsRaw As Worksheet
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("Dati Raw")
    On Error GoTo 0
    If wsRaw Is Nothing Then
        Set wsRaw = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsRaw.Name = "Dati Raw"
    Else
        wsRaw.Cells.Clear
    End If

    Dim headerScritto As Boolean
    headerScritto = False
    Dim rigaCorrente As Long
    rigaCorrente = 1

    ' Trova tutti i CSV nella cartella
    Dim nomeFile As String
    nomeFile = Dir(cartella & "*.csv")
    Dim fileCount As Integer
    fileCount = 0

    Do While nomeFile <> ""
        fileCount = fileCount + 1
        Dim percorso As String
        percorso = cartella & nomeFile

        ' Leggi il CSV
        Dim fileNum As Integer
        fileNum = FreeFile
        Open percorso For Input As #fileNum

        Dim linea As String
        Dim primaRiga As Boolean
        primaRiga = True
        Dim colonne() As String

        Do While Not EOF(fileNum)
            Line Input #fileNum, linea
            linea = Replace(linea, Chr(13), "")

            ' Auto-detect separatore
            Dim sep As String
            If primaRiga Then
                If InStr(linea, ";") > InStr(linea, ",") And _
                   UBound(Split(linea, ";")) > UBound(Split(linea, ",")) Then
                    sep = ";"
                Else
                    sep = ","
                End If
            End If

            Dim campi() As String
            campi = Split(linea, sep)

            If primaRiga Then
                colonne = campi
                If Not headerScritto Then
                    ' Scrivi intestazione + colonna Sorgente
                    Dim c As Integer
                    For c = 0 To UBound(campi)
                        wsRaw.Cells(1, c + 1).Value = campi(c)
                    Next c
                    wsRaw.Cells(1, UBound(campi) + 2).Value = "Sorgente_File"
                    wsRaw.Rows(1).Font.Bold = True
                    wsRaw.Rows(1).Interior.Color = RGB(31, 78, 121)
                    wsRaw.Rows(1).Font.Color = RGB(255, 255, 255)
                    headerScritto = True
                End If
                primaRiga = False
            Else
                ' Scrivi riga dati
                rigaCorrente = rigaCorrente + 1
                Dim k As Integer
                For k = 0 To UBound(campi)
                    Dim val As String
                    val = campi(k)
                    ' Converti numeri
                    If IsNumeric(Replace(val, ",", ".")) Then
                        wsRaw.Cells(rigaCorrente, k + 1).Value = CDbl(Replace(val, ",", "."))
                    Else
                        wsRaw.Cells(rigaCorrente, k + 1).Value = val
                    End If
                Next k
                wsRaw.Cells(rigaCorrente, UBound(campi) + 2).Value = nomeFile
            End If
        Loop

        Close #fileNum
        nomeFile = Dir()
    Loop

    ' Formatta come tabella
    If rigaCorrente > 1 Then
        Dim rng As Range
        Set rng = wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(rigaCorrente, UBound(colonne) + 2))
        wsRaw.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "TabellaRaw"
    End If

    wsRaw.Columns.AutoFit
    MsgBox fileCount & " file CSV importati — " & rigaCorrente - 1 & " righe totali.", vbInformation, "Import CSV"
End Sub

' ============================================================
' CALCOLO STATISTICHE MENSILI
' ============================================================
Sub CalcolaStatistiche()
    Dim wsRaw As Worksheet
    Dim wsStat As Worksheet

    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("Dati Raw")
    On Error GoTo 0
    If wsRaw Is Nothing Then MsgBox "Prima importa i CSV.", vbExclamation: Exit Sub

    ' Crea foglio Statistiche
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Statistiche").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsStat = ThisWorkbook.Sheets.Add(After:=wsRaw)
    wsStat.Name = "Statistiche"

    ' Intestazioni
    Dim headers() As Variant
    headers = Array("Mese", "Misurazioni", _
                    "Vento_TOP92_Avg", "Vento_TOP92_Max", "Vento_TOP92_P50", "Vento_TOP92_P75", "Vento_TOP92_P90", _
                    "Vento_RIF88_Avg", "Vento_RIF70_Avg", "Vento_RIF50_Avg", _
                    "Shear_Alpha_92_50", _
                    "Temp_Avg", "Temp_Max", "Temp_Min", _
                    "Umidita_Avg", "Pressione_Avg", _
                    "Disponibilita_pct")

    Dim h As Integer
    For h = 0 To UBound(headers)
        wsStat.Cells(1, h + 1).Value = headers(h)
    Next h

    ' Stile intestazione
    With wsStat.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Trova indici colonne nel raw
    Dim wsRawLastCol As Long
    wsRawLastCol = wsRaw.Cells(1, wsRaw.Columns.Count).End(xlToLeft).Column

    Dim idxDT As Long, idxW92 As Long, idxW88 As Long, idxW70 As Long, idxW50 As Long
    Dim idxT As Long, idxH As Long, idxP As Long
    idxDT = 0: idxW92 = 0: idxW88 = 0: idxW70 = 0: idxW50 = 0
    idxT = 0: idxH = 0: idxP = 0

    Dim col As Long
    For col = 1 To wsRawLastCol
        Dim hdr As String
        hdr = wsRaw.Cells(1, col).Value
        If hdr = "datetime" Then idxDT = col
        If InStr(1, hdr, "TOP 92") > 0 And InStr(1, hdr, "wind_speed") > 0 And InStr(1, hdr, "Avg") > 0 Then idxW92 = col
        If InStr(1, hdr, "RIF 88") > 0 And InStr(1, hdr, "wind_speed") > 0 And InStr(1, hdr, "Avg") > 0 Then idxW88 = col
        If InStr(1, hdr, "RIF 70") > 0 And InStr(1, hdr, "wind_speed") > 0 And InStr(1, hdr, "Avg") > 0 Then idxW70 = col
        If InStr(1, hdr, "RIF 50") > 0 And InStr(1, hdr, "wind_speed") > 0 And InStr(1, hdr, "Avg") > 0 Then idxW50 = col
        If InStr(1, hdr, "temperature") > 0 And InStr(1, hdr, "Avg") > 0 Then idxT = col
        If InStr(1, hdr, "humidity") > 0 And InStr(1, hdr, "Avg") > 0 Then idxH = col
        If InStr(1, hdr, "air_pressure") > 0 And InStr(1, hdr, "Avg") > 0 Then idxP = col
    Next col

    ' Raggruppa per mese usando dizionario
    Dim ultimaRiga As Long
    ultimaRiga = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' Raccogli mesi unici
    Dim mesi() As String
    ReDim mesi(0)
    Dim nMesi As Integer
    nMesi = 0

    Dim r As Long
    For r = 2 To ultimaRiga
        Dim dtStr As String
        dtStr = CStr(wsRaw.Cells(r, idxDT).Value)
        If Len(dtStr) >= 7 Then
            Dim meseKey As String
            meseKey = Left(dtStr, 7)
            ' Controlla se già presente
            Dim trovato As Boolean
            trovato = False
            Dim m As Integer
            For m = 0 To nMesi - 1
                If mesi(m) = meseKey Then trovato = True: Exit For
            Next m
            If Not trovato Then
                ReDim Preserve mesi(nMesi)
                mesi(nMesi) = meseKey
                nMesi = nMesi + 1
            End If
        End If
    Next r

    ' Ordina mesi
    Dim i As Integer, j As Integer, temp As String
    For i = 0 To nMesi - 2
        For j = i + 1 To nMesi - 1
            If mesi(i) > mesi(j) Then
                temp = mesi(i): mesi(i) = mesi(j): mesi(j) = temp
            End If
        Next j
    Next i

    ' Per ogni mese calcola statistiche
    Dim rigaStat As Long
    rigaStat = 2

    For m = 0 To nMesi - 1
        Dim mk As String
        mk = mesi(m)

        Dim cntR As Long, sumW92 As Double, maxW92 As Double
        Dim sumW88 As Double, sumW70 As Double, sumW50 As Double
        Dim sumT As Double, maxT As Double, minT As Double
        Dim sumH As Double, sumP As Double
        Dim cntT As Long, cntP As Long
        Dim w92Vals() As Double
        ReDim w92Vals(0)
        Dim nW92 As Long
        nW92 = 0

        cntR = 0: sumW92 = 0: maxW92 = 0
        sumW88 = 0: sumW70 = 0: sumW50 = 0
        sumT = 0: maxT = -999: minT = 999
        sumH = 0: sumP = 0
        cntT = 0: cntP = 0

        For r = 2 To ultimaRiga
            dtStr = CStr(wsRaw.Cells(r, idxDT).Value)
            If Left(dtStr, 7) = mk Then
                cntR = cntR + 1
                ' Vento 92
                If idxW92 > 0 Then
                    Dim vW92 As Double
                    vW92 = 0
                    On Error Resume Next
                    vW92 = CDbl(wsRaw.Cells(r, idxW92).Value)
                    On Error GoTo 0
                    If vW92 > 0 Then
                        sumW92 = sumW92 + vW92
                        If vW92 > maxW92 Then maxW92 = vW92
                        ReDim Preserve w92Vals(nW92)
                        w92Vals(nW92) = vW92
                        nW92 = nW92 + 1
                    End If
                End If
                ' Altezze inferiori
                If idxW88 > 0 Then On Error Resume Next: sumW88 = sumW88 + CDbl(wsRaw.Cells(r, idxW88).Value): On Error GoTo 0
                If idxW70 > 0 Then On Error Resume Next: sumW70 = sumW70 + CDbl(wsRaw.Cells(r, idxW70).Value): On Error GoTo 0
                If idxW50 > 0 Then On Error Resume Next: sumW50 = sumW50 + CDbl(wsRaw.Cells(r, idxW50).Value): On Error GoTo 0
                ' Temperatura
                If idxT > 0 Then
                    Dim vT As Double
                    vT = 0
                    On Error Resume Next
                    vT = CDbl(wsRaw.Cells(r, idxT).Value)
                    On Error GoTo 0
                    If vT > -10 And vT < 60 Then
                        sumT = sumT + vT: cntT = cntT + 1
                        If vT > maxT Then maxT = vT
                        If vT < minT Then minT = vT
                    End If
                End If
                ' Umidità e pressione
                If idxH > 0 Then On Error Resume Next: sumH = sumH + CDbl(wsRaw.Cells(r, idxH).Value): On Error GoTo 0
                If idxP > 0 Then
                    Dim vP As Double
                    vP = 0
                    On Error Resume Next
                    vP = CDbl(wsRaw.Cells(r, idxP).Value)
                    On Error GoTo 0
                    If vP > 900 And vP < 1100 Then sumP = sumP + vP: cntP = cntP + 1
                End If
            End If
        Next r

        ' Calcola percentili P50 P75 P90 (eccedenza)
        Dim p50 As Double, p75 As Double, p90 As Double
        p50 = 0: p75 = 0: p90 = 0
        If nW92 > 0 Then
            ' Ordina
            For i = 0 To nW92 - 2
                For j = i + 1 To nW92 - 1
                    If w92Vals(i) > w92Vals(j) Then
                        Dim tmp As Double
                        tmp = w92Vals(i): w92Vals(i) = w92Vals(j): w92Vals(j) = tmp
                    End If
                Next j
            Next i
            ' Eccedenza: P50=50° percentile, P75=25° pct, P90=10° pct
            p50 = w92Vals(CLng((nW92 - 1) * 0.5))
            p75 = w92Vals(CLng((nW92 - 1) * 0.25))
            p90 = w92Vals(CLng((nW92 - 1) * 0.1))
        End If

        ' Shear alpha (92m vs 50m)
        Dim shear As Double
        shear = 0
        If sumW92 > 0 And sumW50 > 0 And cntR > 0 Then
            Dim avgW92 As Double, avgW50 As Double
            avgW92 = sumW92 / cntR
            avgW50 = sumW50 / cntR
            If avgW92 > 0 And avgW50 > 0 Then
                shear = Log(avgW92 / avgW50) / Log(92 / 50)
            End If
        End If

        ' Scrivi riga statistiche
        Dim attesi As Long
        ' 10-min intervals: stima giorni nel mese * 24h * 6
        Dim giorniMese As Integer
        giorniMese = 30
        If Right(mk, 2) = "03" Or Right(mk, 2) = "05" Or Right(mk, 2) = "07" Or _
           Right(mk, 2) = "08" Or Right(mk, 2) = "10" Or Right(mk, 2) = "12" Then giorniMese = 31
        If Right(mk, 2) = "02" Then giorniMese = 28
        attesi = giorniMese * 144

        Dim avail As Double
        If attesi > 0 Then avail = WorksheetFunction.Min(100, Round(cntR / attesi * 100, 1)) Else avail = 0

        With wsStat
            .Cells(rigaStat, 1).Value = mk
            .Cells(rigaStat, 2).Value = cntR
            .Cells(rigaStat, 3).Value = IIf(cntR > 0, Round(sumW92 / cntR, 3), 0)
            .Cells(rigaStat, 4).Value = Round(maxW92, 3)
            .Cells(rigaStat, 5).Value = Round(p50, 3)
            .Cells(rigaStat, 6).Value = Round(p75, 3)
            .Cells(rigaStat, 7).Value = Round(p90, 3)
            .Cells(rigaStat, 8).Value = IIf(cntR > 0, Round(sumW88 / cntR, 3), 0)
            .Cells(rigaStat, 9).Value = IIf(cntR > 0, Round(sumW70 / cntR, 3), 0)
            .Cells(rigaStat, 10).Value = IIf(cntR > 0, Round(sumW50 / cntR, 3), 0)
            .Cells(rigaStat, 11).Value = Round(shear, 4)
            .Cells(rigaStat, 12).Value = IIf(cntT > 0, Round(sumT / cntT, 2), 0)
            .Cells(rigaStat, 13).Value = IIf(maxT > -999, Round(maxT, 2), 0)
            .Cells(rigaStat, 14).Value = IIf(minT < 999, Round(minT, 2), 0)
            .Cells(rigaStat, 15).Value = IIf(cntR > 0, Round(sumH / cntR, 1), 0)
            .Cells(rigaStat, 16).Value = IIf(cntP > 0, Round(sumP / cntP, 1), 0)
            .Cells(rigaStat, 17).Value = avail
        End With

        ' Colora righe alternate
        If rigaStat Mod 2 = 0 Then
            wsStat.Rows(rigaStat).Interior.Color = RGB(214, 228, 240)
        End If

        rigaStat = rigaStat + 1
    Next m

    wsStat.Columns.AutoFit
End Sub

' ============================================================
' CREA DASHBOARD CON GRAFICI
' ============================================================
Sub CreaDashboard()
    Dim wsStat As Worksheet
    Dim wsDash As Worksheet

    On Error Resume Next
    Set wsStat = ThisWorkbook.Sheets("Statistiche")
    On Error GoTo 0
    If wsStat Is Nothing Then MsgBox "Prima calcola le statistiche.", vbExclamation: Exit Sub

    ' Crea foglio Dashboard
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Dashboard").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsDash = ThisWorkbook.Sheets.Add(After:=wsStat)
    wsDash.Name = "Dashboard"
    wsDash.Tab.Color = RGB(31, 78, 121)

    ' Titolo
    With wsDash.Range("A1:P1")
        .Merge
        .Value = "Dashboard Meteo — Analisi Stazione Anemometrica"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Data ultimo aggiornamento
    With wsDash.Range("A2:P2")
        .Merge
        .Value = "Aggiornato: " & Format(Now, "dd/mm/yyyy hh:mm")
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 117, 182)
        .HorizontalAlignment = xlCenter
    End With

    ' Calcola range dati statistiche
    Dim ultimaRigaStat As Long
    ultimaRigaStat = wsStat.Cells(wsStat.Rows.Count, 1).End(xlUp).Row

    ' ---- GRAFICO 1: Vento medio per altezza (grouped bar) ----
    Dim ch1 As ChartObject
    Set ch1 = wsDash.ChartObjects.Add(Left:=10, Top:=60, Width:=500, Height:=280)
    With ch1.Chart
        .ChartType = xlColumnClustered
        .SetSourceData wsStat.Range("A1:A" & ultimaRigaStat)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "TOP 92m"
        .SeriesCollection(1).Values = wsStat.Range("C2:C" & ultimaRigaStat)
        .SeriesCollection(1).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(1).Interior.Color = RGB(31, 78, 121)
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "RIF 88m"
        .SeriesCollection(2).Values = wsStat.Range("H2:H" & ultimaRigaStat)
        .SeriesCollection(2).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(2).Interior.Color = RGB(46, 117, 182)
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "RIF 70m"
        .SeriesCollection(3).Values = wsStat.Range("I2:I" & ultimaRigaStat)
        .SeriesCollection(3).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(3).Interior.Color = RGB(29, 158, 117)
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Name = "RIF 50m"
        .SeriesCollection(4).Values = wsStat.Range("J2:J" & ultimaRigaStat)
        .SeriesCollection(4).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(4).Interior.Color = RGB(186, 117, 23)
        .HasTitle = True
        .ChartTitle.Text = "Velocità Vento Media per Altezza (m/s)"
        .ChartTitle.Font.Size = 12
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "m/s"
        .PlotArea.Interior.Color = RGB(248, 251, 255)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With

    ' ---- GRAFICO 2: P50 P75 P90 ----
    Dim ch2 As ChartObject
    Set ch2 = wsDash.ChartObjects.Add(Left:=520, Top:=60, Width:=500, Height:=280)
    With ch2.Chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "P50 (50% eccedenza)"
        .SeriesCollection(1).Values = wsStat.Range("E2:E" & ultimaRigaStat)
        .SeriesCollection(1).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(1).Border.Color = RGB(29, 158, 117)
        .SeriesCollection(1).Border.Weight = xlMedium
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "P75 (75% eccedenza)"
        .SeriesCollection(2).Values = wsStat.Range("F2:F" & ultimaRigaStat)
        .SeriesCollection(2).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(2).Border.Color = RGB(212, 160, 23)
        .SeriesCollection(2).Border.Weight = xlMedium
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "P90 (90% eccedenza)"
        .SeriesCollection(3).Values = wsStat.Range("G2:G" & ultimaRigaStat)
        .SeriesCollection(3).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(3).Border.Color = RGB(186, 117, 23)
        .SeriesCollection(3).Border.Weight = xlMedium
        .HasTitle = True
        .ChartTitle.Text = "P50 / P75 / P90 — Velocità Vento (m/s)"
        .ChartTitle.Font.Size = 12
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "m/s"
        .PlotArea.Interior.Color = RGB(248, 251, 255)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With

    ' ---- GRAFICO 3: Temperatura ----
    Dim ch3 As ChartObject
    Set ch3 = wsDash.ChartObjects.Add(Left:=10, Top:=360, Width:=500, Height:=250)
    With ch3.Chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Media"
        .SeriesCollection(1).Values = wsStat.Range("L2:L" & ultimaRigaStat)
        .SeriesCollection(1).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(1).Border.Color = RGB(216, 90, 48)
        .SeriesCollection(1).Border.Weight = xlMedium
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Max"
        .SeriesCollection(2).Values = wsStat.Range("M2:M" & ultimaRigaStat)
        .SeriesCollection(2).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(2).Border.Color = RGB(192, 57, 43)
        .SeriesCollection(2).Border.LineStyle = xlDash
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "Min"
        .SeriesCollection(3).Values = wsStat.Range("N2:N" & ultimaRigaStat)
        .SeriesCollection(3).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .SeriesCollection(3).Border.Color = RGB(41, 128, 185)
        .SeriesCollection(3).Border.LineStyle = xlDash
        .HasTitle = True
        .ChartTitle.Text = "Temperatura Mensile (°C)"
        .ChartTitle.Font.Size = 12
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "°C"
        .PlotArea.Interior.Color = RGB(248, 251, 255)
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With

    ' ---- GRAFICO 4: Disponibilità ----
    Dim ch4 As ChartObject
    Set ch4 = wsDash.ChartObjects.Add(Left:=520, Top:=360, Width:=500, Height:=250)
    With ch4.Chart
        .ChartType = xlColumnClustered
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Disponibilità %"
        .SeriesCollection(1).Values = wsStat.Range("Q2:Q" & ultimaRigaStat)
        .SeriesCollection(1).XValues = wsStat.Range("A2:A" & ultimaRigaStat)
        .HasTitle = True
        .ChartTitle.Text = "Disponibilità del Dato (%)"
        .ChartTitle.Font.Size = 12
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 100
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "%"
        .PlotArea.Interior.Color = RGB(248, 251, 255)
        .HasLegend = False
        ' Colora barre in verde
        .SeriesCollection(1).Interior.Color = RGB(29, 158, 117)
    End With

    ' ---- KPI CARDS in alto ----
    Call AggiungiKPI(wsDash, wsStat, ultimaRigaStat)
End Sub

' ============================================================
' KPI CARDS
' ============================================================
Sub AggiungiKPI(wsDash As Worksheet, wsStat As Worksheet, ultimaRiga As Long)
    Dim kpiRow As Long
    kpiRow = 630 ' pixel offset — adatta se necessario

    ' Media annuale vento
    Dim avgVento As Double
    avgVento = WorksheetFunction.Average(wsStat.Range("C2:C" & ultimaRiga))

    ' Disponibilità media
    Dim avgAvail As Double
    avgAvail = WorksheetFunction.Average(wsStat.Range("Q2:Q" & ultimaRiga))

    ' P50 medio
    Dim avgP50 As Double
    avgP50 = WorksheetFunction.Average(wsStat.Range("E2:E" & ultimaRiga))

    ' Shear medio
    Dim avgShear As Double
    avgShear = WorksheetFunction.Average(wsStat.Range("K2:K" & ultimaRiga))

    ' Scrivi KPI nel foglio (sotto i grafici)
    Dim startRow As Long
    startRow = 40

    Dim kpiData(3, 2) As String
    kpiData(0, 0) = "Vento Medio TOP 92m"
    kpiData(0, 1) = Format(avgVento, "0.00") & " m/s"
    kpiData(0, 2) = "Media su tutti i mesi"
    kpiData(1, 0) = "Disponibilità Media"
    kpiData(1, 1) = Format(avgAvail, "0.0") & " %"
    kpiData(1, 2) = "Dati disponibili / attesi"
    kpiData(2, 0) = "P50 Medio Annuale"
    kpiData(2, 1) = Format(avgP50, "0.00") & " m/s"
    kpiData(2, 2) = "Eccedenza 50% del tempo"
    kpiData(3, 0) = "Wind Shear α Medio"
    kpiData(3, 1) = Format(avgShear, "0.000")
    kpiData(3, 2) = "Legge della potenza 50-92m"

    Dim kpiColors(3) As Long
    kpiColors(0) = RGB(31, 78, 121)
    kpiColors(1) = RGB(29, 158, 117)
    kpiColors(2) = RGB(212, 160, 23)
    kpiColors(3) = RGB(142, 68, 173)

    Dim k As Integer
    For k = 0 To 3
        Dim startCol As Long
        startCol = k * 4 + 1
        With wsDash.Range(wsDash.Cells(startRow, startCol), wsDash.Cells(startRow + 3, startCol + 3))
            .Merge
            .Interior.Color = kpiColors(k)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 10
            .Value = kpiData(k, 0) & vbCrLf & kpiData(k, 1) & vbCrLf & kpiData(k, 2)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next k
End Sub

' ============================================================
' ESPORTA JSON PER STREAMLIT
' ============================================================
Sub EsportaJSON()
    Dim wsStat As Worksheet
    On Error Resume Next
    Set wsStat = ThisWorkbook.Sheets("Statistiche")
    On Error GoTo 0
    If wsStat Is Nothing Then Exit Sub

    Dim ultimaRiga As Long
    ultimaRiga = wsStat.Cells(wsStat.Rows.Count, 1).End(xlUp).Row

    Dim json As String
    json = "[" & vbCrLf

    Dim r As Long
    For r = 2 To ultimaRiga
        json = json & "  {" & vbCrLf
        json = json & "    ""mese"": """ & wsStat.Cells(r, 1).Value & """," & vbCrLf
        json = json & "    ""misurazioni"": " & wsStat.Cells(r, 2).Value & "," & vbCrLf
        json = json & "    ""vento_top92_avg"": " & wsStat.Cells(r, 3).Value & "," & vbCrLf
        json = json & "    ""vento_top92_max"": " & wsStat.Cells(r, 4).Value & "," & vbCrLf
        json = json & "    ""p50"": " & wsStat.Cells(r, 5).Value & "," & vbCrLf
        json = json & "    ""p75"": " & wsStat.Cells(r, 6).Value & "," & vbCrLf
        json = json & "    ""p90"": " & wsStat.Cells(r, 7).Value & "," & vbCrLf
        json = json & "    ""vento_88m"": " & wsStat.Cells(r, 8).Value & "," & vbCrLf
        json = json & "    ""vento_70m"": " & wsStat.Cells(r, 9).Value & "," & vbCrLf
        json = json & "    ""vento_50m"": " & wsStat.Cells(r, 10).Value & "," & vbCrLf
        json = json & "    ""shear_alpha"": " & wsStat.Cells(r, 11).Value & "," & vbCrLf
        json = json & "    ""temp_avg"": " & wsStat.Cells(r, 12).Value & "," & vbCrLf
        json = json & "    ""temp_max"": " & wsStat.Cells(r, 13).Value & "," & vbCrLf
        json = json & "    ""temp_min"": " & wsStat.Cells(r, 14).Value & "," & vbCrLf
        json = json & "    ""umidita_avg"": " & wsStat.Cells(r, 15).Value & "," & vbCrLf
        json = json & "    ""pressione_avg"": " & wsStat.Cells(r, 16).Value & "," & vbCrLf
        json = json & "    ""disponibilita_pct"": " & wsStat.Cells(r, 17).Value & vbCrLf
        json = json & "  }"
        If r < ultimaRiga Then json = json & ","
        json = json & vbCrLf
    Next r

    json = json & "]"

    ' Salva il file JSON nella stessa cartella del workbook
    Dim percorso As String
    percorso = ThisWorkbook.Path & "\export_streamlit.json"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open percorso For Output As #fileNum
    Print #fileNum, json
    Close #fileNum

    MsgBox "JSON esportato in: " & percorso, vbInformation, "Export JSON"
End Sub
