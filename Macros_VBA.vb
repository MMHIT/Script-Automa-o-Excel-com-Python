'Macros VBA
                   
                    Sub Copia_Arquivo_Banco ()
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

                   
                    Sub Formata_Liquid ()
                    
                        ' limpa Macro
                        ' Formata como texto
                        '
                            Columns("A:Z").Select
                            Selection.NumberFormat = "@"
                            Range("A1").Select


                    ' Formata Macro
                    ' Move e Converte
                    '
                        Columns("A:Y").Select
                        Columns("A:Y").EntireColumn.AutoFit
                    '    Range("G1").Select
                    '   Selection.Cut Destination:=Range("H1")
                    '    Range("I1").Select
                    '    Selection.Cut Destination:=Range("J1")
                    '    Range("G:G,I:I").Select
                    '    Range("I1").Activate
                    '    Selection.Delete Shift:=xlToLeft
                    '    Columns("J:J").Select
                    '    Selection.Delete Shift:=xlToLeft
                    '    ActiveWindow.SmallScroll ToRight:=8
                        Selection.Replace What:="LIQUIDAÇÃO", Replacement:="LIQUIDACAO", LookAt:= _
                            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                        Selection.Replace What:="NÃO", Replacement:="NAO", LookAt:= _
                            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                    '    Range("I2").Select
                    '    Cells.Replace What:="LIQUIDAÇÃO", Replacement:="LIQUIDACAO", LookAt:= _
                    '        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    '        ReplaceFormat:=False
                    '    Range("I2").Select
                        Cells.Replace What:="COMPENSAÇÃO", Replacement:="COMPENSACAO", LookAt:= _
                            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                    '    Range("R2").Select
                        Cells.Replace What:="AUTOMÁTICA", Replacement:="AUTOMATICA", LookAt:= _
                            xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                        Cells.Replace What:="R$", Replacement:="", LookAt:=xlPart, SearchOrder _
                        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        'ChDir "C:\Users\pvmatheus\Desktop"
                        'ActiveWorkbook.SaveAs Filename:="C:\Users\pvmatheus\Desktop\Pasta1.xlsx", _
                        'FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                    End Sub


                    Sub Formata_Entrada ()
                    
                        ' Limpa Macro
                        ' Limpa tabela
                        '
                            Columns("A:AK").Select
                            Selection.NumberFormat = "@"
                            Range("A1").Select


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
                        Columns("AA:AD").Select
                        Columns("AA:AD").EntireColumn.AutoFit
                        Range("A1").Select
                        'ChDir "C:\Users\pvmatheus\Desktop"
                        'ActiveWorkbook.SaveAs Filename:="C:\Users\pvmatheus\Desktop\Pasta1.xlsx", _
                            'FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                    End Sub


                    Sub Exporta ()

                    ' Tranzita entre o excel e o arquivo texto

                    Dim arquivo As String, texto As String, linhaTexto As String
                    Dim chegada As String, partida As String

                    arquivo = "C:\Users\pvmatheus\Desktop\BotExcelSicoob\Arquivos_txt\entrada_3298.txt"

                    Open arquivo For Input As #1

                    Do Until EOF(1)

                        Line Input #1, linhaTexto

                        texto = texto & linhaTexto

                    Loop

                    Close #1

                    partida = InStr(texto, “horarioPartida: “)

                    chegada = InStr(texto, “horarioChegada: “)

                    Range(“A1”).Value = Mid(texto, partida + 16, 5)

                    Range(“A2”).Value = Mid(texto, chegada + 16, 5)
   
                    End Sub

                    'Sub limpa_liquid ()
                    '
                    ' limpa Macro
                    ' Formata como texto
                    '
                    ' Atalho do teclado: Ctrl+l
                    '
                    '    Columns("A:Z").Select
                    '    Selection.NumberFormat = "@"
                    '    Range("A1").Select
                    'End Sub

                    'Sub Limpa_entrada ()
                    '
                    ' Limpa Macro
                    ' Limpa tabela
                    '
                    ' Atalho do teclado: Ctrl+l
                    '
                    '    Columns("A:AK").Select
                    '    Selection.NumberFormat = "@"
                    '    Range("A1").Select
                    'End Sub