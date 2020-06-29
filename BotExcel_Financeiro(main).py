# Bot Financeiro inserção de dados.
    from datetime       import date
    from openpyxl       import Workbook
    from openpyxl.utils import get_column_letter
    from openpyxl       import load_workbook

# Data Atualizada
    data_atual = date.today()
    data_em_texto = data_atual.strftime(‘%d-%m-%Y’)

    #Disparo de e-mail ao Banco
    
    #Pegando o arquivo do banco



    #lendo o arquivo do banco

        #    wb = load_workbook('Entradas_3298.xlsx') # abrindo o Workbook test.xlsx
        #    ws = wb['Page 1'] # selecionando a planilha 'Plan1' dentro do Workbook test.xlsx

        #    for line in ws:  # iterando em todas as linhas da 'Plan1'
        #       print line[3] # print a primeira célula da linha


        #    wb = load_workbook('Liquidacoes_3298.xlsx') # abrindo o Workbook test.xlsx
        #    ws = wb['Page 1'] # selecionando a planilha 'Plan1' dentro do Workbook test.xlsx

        #    for line in ws:  # iterando em todas as linhas da 'Plan1'
        #        print line[3] # print a primeira célula da linha


        #    wb = load_workbook('Entradas_15423.xlsx') # abrindo o Workbook test.xlsx
        #    ws = wb['Page 1'] # selecionando a planilha 'Plan1' dentro do Workbook test.xlsx

        #    for line in ws:  # iterando em todas as linhas da 'Plan1'
        #        print line[3] # print a primeira célula da linha
                

        #    wb = load_workbook('Liquidacoes_15423.xlsx') # abrindo o Workbook test.xlsx
        #    ws = wb['Page 1'] # selecionando a planilha 'Plan1' dentro do Workbook test.xlsx

        #    for line in ws:  # iterando em todas as linhas da 'Plan1'
        #        print line[3] # print a primeira célula da linha



    #Processando Arquivo            
    
         #Cria uma pasta de trabalho no excel(arquivo) em excel
          from openpyxl import Workbook

            Goiania_entrada = Workbook()
                #Ativa a planilha no excel
                    planilha1 = Goiania_entrada.active
                    #Nomeia a planilha criada
                        planilha1.title = "3298_entrada_" + data_em_texto

            arquivo_excel.save("3298_entrada_" + data_em_texto ".xlsx") #Salva arquivo com nome sem a data referente

            #Copiando arquivos do Escel sicoob para o novo arquivo
                original = arquivo_excel.get_sheet_by_name('Entradas_3298.xlsx')
                copia = arquivo_excel.copy_worksheet(copia)
                arquivo_excel.save('3298_entrada_' + data_em_texto '.xlsx')

            Goiania_liquidação = Workbook()
                #Ativa a planilha no excel
                    planilha1 = Goiania_liquidação.active
                    #Nomeia a planilha criada
                        planilha1.title = "3298_liquid_" + data_em_texto

            arquivo_excel.save("3298_liquid_" + data_em_texto ".xlsx") #Salva arquivo com nome sem data referente

            #Copiando arquivos do Escel sicoob para o novo arquivo
                original = arquivo_excel.get_sheet_by_name('Liquidacoes_3298.xlsx')
                copia = arquivo_excel.copy_worksheet(copia)
                arquivo_excel.save('3298_liquid_' + data_em_texto '.xlsx')


                    #2º exemplo de nomeação (******planilha2 = Goiania_entrada.create_sheet("Nova entrada_27-05-29_05_2020_3")*********)
                    #3º exemplo de nomeação em posição definida (****planilha2 = Goiania_liquidação.create_sheet("Nova entrada_27-05-29_05_2020_3", 0)*****)
         
    #Formatando o arquivo
            wb = Workbook()

            dest_filename = '3298_entrada_.xlsx'
            
            ws1 = wb.active
            ws1.title = "3298_entrada_"
            
            for row in range(1, 40):
                ws1.append(range(600))
                ws2 = wb.create_sheet(title="Pi")
                ws2['F5'] = 3.14
                ws3 = wb.create_sheet(title="Data")
            for row in range(10, 20):
                for col in range(27, 54):
                   _ = ws3.cell(column=col, row=row, value="{0}".format(get_column_letter(col)))
                     print(ws3['AA10'].value)
                    AA
                    wb.save(filename = dest_filename)

    #Salva em CSV


       

          