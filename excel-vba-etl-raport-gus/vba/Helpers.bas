Attribute VB_Name = "Module1"
Option Explicit

' ======================================================
' Mapowanie wyboru u�ytkownika na dok�adne nazwy wska�nik�w GUS
' ======================================================
Function GetWskaznik() As String
    Select Case Sheets("START").Range("C3").Value
        Case "Nominalne wynagrodzenie � z�"
            GetWskaznik = "Przeci�tne miesi�czne nominalne wynagrodzenie brutto w sektorze przedsi�biorstw"
        Case "Realne wynagrodzenie � r/r"
            GetWskaznik = "Przeci�tne miesi�czne realne wynagrodzenie brutto w sektorze przedsi�biorstw"
        Case "Emerytury � z�"
            GetWskaznik = "Przeci�tna miesi�czna nominalna emerytura i renta brutto z pozarolniczego systemu ubezpiecze� spo�ecznych"
        Case Else
            GetWskaznik = ""
    End Select
End Function
