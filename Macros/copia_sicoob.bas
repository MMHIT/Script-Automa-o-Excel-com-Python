Attribute VB_Name = "Copia_sicoob"
Option Explicit
Sub Copia_Arquivo_Banco()
 '
 ' Copia Macro
 ' Copia e formata texto
 '
 ' Atalho do teclado: Ctrl+j
 '
 Range("A4").Select
 Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
 Selection.NumberFormat = "@"
 Selection.Copy
 End Sub

