Attribute VB_Name = "Module1"
Option Explicit

' ======================================================
' Import danych CSV z kodowaniem UTF-8 (GUS)
' Pomini�cie QueryTables � r�czne parsowanie pliku
' ======================================================
Sub ImportCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim stm As Object
    Dim content As String
    Dim lines As Variant
    Dim fields As Variant
    Dim r As Long, c As Long

    Set ws = Sheets("DANE_RAW")
    ws.Cells.Clear

    ' Wyb�r pliku CSV przez u�ytkownika
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Wybierz plik CSV"
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        If .Show = False Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' Odczyt pliku jako UTF-8 (ADODB.Stream)
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 ' text
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    content = stm.ReadText
    stm.Close

    ' Podzia� na linie i pola (separator ;)
    lines = Split(content, vbLf)

    For r = LBound(lines) To UBound(lines)
        If Trim(lines(r)) <> "" Then
            fields = Split(lines(r), ";")
            For c = LBound(fields) To UBound(fields)
                ' Usuni�cie cudzys�ow�w z p�l tekstowych
                ws.Cells(r + 1, c + 1).Value = Replace(fields(c), """", "")
            Next c
        End If
    Next r

    ws.Columns.AutoFit
    Sheets("START").Range("C5").Value = "Dane wczytane"
End Sub
