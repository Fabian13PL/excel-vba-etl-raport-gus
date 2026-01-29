Attribute VB_Name = "Module1"
Option Explicit

' ======================================================
' Generowanie wykresu na podstawie tabeli przestawnej
' ======================================================
Sub CreateChart()
    Dim ch As Chart
    
    Sheets("RAPORT").Cells.Clear
    
    Set ch = Charts.Add
    ch.ChartType = xlLine
    ch.SetSourceData Source:=Sheets("PIVOT").Range("A3").CurrentRegion
    ch.Location Where:=xlLocationAsObject, Name:="RAPORT"
End Sub
