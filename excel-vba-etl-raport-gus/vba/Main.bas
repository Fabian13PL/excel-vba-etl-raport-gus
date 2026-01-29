Attribute VB_Name = "Module1"
Option Explicit

' Wrapper do przycisku �Importuj dane�
Sub Import()
    Call ImportCSV
End Sub

' Wrapper do przycisku �Generuj raport�
Sub Generate()
    Call GenerateReport
End Sub

' ======================================================
' Pe�ny pipeline: transformacja � pivot � wykres
' ======================================================
Sub GenerateReport()
    Application.ScreenUpdating = False

    Call BuildDANELong
    Call CreatePivot
    Call CreateChart

    Sheets("START").Range("C5").Value = "Raport wygenerowany"
    Application.ScreenUpdating = True
End Sub
