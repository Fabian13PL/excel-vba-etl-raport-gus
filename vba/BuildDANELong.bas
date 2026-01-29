Attribute VB_Name = "Module1"
Option Explicit

' ======================================================
' Transformacja danych GUS z uk�adu szerokiego (wide)
' do postaci analitycznej (long / tidy data)
'
' Efekt:
' Wska�nik | Jednostka | Rok | Miesi�c | Warto��
' ======================================================
Sub BuildDANELong()
    Dim src As Worksheet, tgt As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim outRow As Long
    Dim rok As Long, miesiac As Long
    Dim wartosc As String

    Set src = Sheets("DANE_RAW")
    Set tgt = Sheets("DANE_LONG")

    tgt.Cells.Clear
    tgt.Range("A1:E1").Value = Array("Wska�nik", "Jednostka", "Rok", "Miesi�c", "Warto��")

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column

    outRow = 2

    ' Dane:
    ' wiersze od 3 � obserwacje
    ' kolumny od 4 � warto�ci miesi�czne
    For r = 3 To lastRow
        For c = 4 To lastCol

            wartosc = Trim(src.Cells(r, c).Value)

            ' Pomijanie brak�w danych
            If wartosc <> "" And wartosc <> "." Then

                ' Normalizacja separatora dziesi�tnego
                wartosc = Replace(wartosc, ",", ".")
                wartosc = Trim(wartosc)

                ' Val() odporne na �brudne� dane tekstowe
                If Val(wartosc) <> 0 Then
                    rok = CLng(src.Cells(1, c).Value)
                    miesiac = CLng(src.Cells(2, c).Value)

                    tgt.Cells(outRow, 1).Value = src.Cells(r, 1).Value   ' Wska�nik
                    tgt.Cells(outRow, 2).Value = src.Cells(r, 3).Value   ' Jednostka
                    tgt.Cells(outRow, 3).Value = rok
                    tgt.Cells(outRow, 4).Value = miesiac
                    tgt.Cells(outRow, 5).Value = Val(wartosc)

                    outRow = outRow + 1
                End If
            End If

        Next c
    Next r

    tgt.Columns.AutoFit
End Sub
