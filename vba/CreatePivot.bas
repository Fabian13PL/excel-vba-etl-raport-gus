Attribute VB_Name = "Module1"
Option Explicit

' ======================================================
' Budowa tabeli przestawnej na danych LONG
' ======================================================
Sub CreatePivot()
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim src As Worksheet
    Dim lastRow As Long, lastCol As Long

    Set ws = Sheets("PIVOT")
    Set src = Sheets("DANE_LONG")

    ws.Cells.Clear

    lastRow = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column

    Set pc = ThisWorkbook.PivotCaches.Create( _
        xlDatabase, _
        src.Range(src.Cells(1, 1), src.Cells(lastRow, lastCol)))

    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A3"), _
        TableName:="PivotWynagrodzenia")

    With pt
        .PivotFields("Rok").Orientation = xlRowField
        .PivotFields("Miesi�c").Orientation = xlColumnField
        .PivotFields("Wska�nik").Orientation = xlPageField
        .AddDataField .PivotFields("Warto��"), "�rednia", xlAverage

        ' Ustawienie wska�nika na podstawie wyboru u�ytkownika
        If GetWskaznik <> "" Then
            .PivotFields("Wska�nik").CurrentPage = GetWskaznik
        End If
    End With
End Sub
