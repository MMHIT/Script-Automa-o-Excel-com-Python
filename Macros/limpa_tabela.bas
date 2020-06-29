Attribute VB_Name = "limpa_tabela"
Option Explicit
Sub Limpa()

                        ' limpa Macro
                        ' Formata como texto
                        '
                            Columns("A:Z").Select
                            Selection.NumberFormat = "@"
                            Range("A1").Select
End Sub
