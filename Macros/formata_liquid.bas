Attribute VB_Name = "formata_liquid"
 Option Explicit
  Sub formata_liquid()
  
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

