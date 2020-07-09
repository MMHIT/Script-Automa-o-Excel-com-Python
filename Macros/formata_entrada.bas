Attribute VB_Name = "formata_entrada"
Sub formata_entrada()

                    ' Formata Macro
                    ' Formata tabela Remove os cifrões
                    '
                    ' Atalho do teclado: Ctrl+h
                    '
                        Columns("A:AK").Select
                        Columns("A:AK").EntireColumn.AutoFit
                    '    Range("F1").Select
                    '    Selection.Cut Destination:=Range("G1")
                    '    Range("G1").Select
                    '    Range("AF1").Select
                    '    Selection.Cut Destination:=Range("AG1")
                    '    Range("AH1").Select
                    '    Selection.Cut Destination:=Range("AI1")
                    '    Range("AJ1").Select
                    '    Selection.Cut Destination:=Range("AK1")
                    '    Range("AK1").Select
                    '    Range("E:E,F:F,H:H,M:M,AE:AE,AF:AF,AJ:AJ").Select
                    '    Range("AJ1").Activate
                    '    Selection.Delete Shift:=xlToLeft
                        Cells.Replace What:="R$ ", Replacement:="", LookAt:=xlPart, SearchOrder _
                            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                    '    Range("AB:AB").Select
                    '    Selection.Delete Shift:=xlToLeft
                    '    Columns("AA:AD").Select
                    '    Columns("AA:AD").EntireColumn.AutoFit
                        Range("A1").Select
                        'ChDir "C:\Users\pvmatheus\Desktop"
                        'ActiveWorkbook.SaveAs Filename:="C:\Users\pvmatheus\Desktop\Pasta1.xlsx", _
                            'FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

