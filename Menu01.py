import fundamentus
import openpyxl
import pandas as pd
import datetime
import sqlite3
import csv
import json
import requests
import os
from io import StringIO
from print_color import print
#-------------------------------------------------------------------------------------------------------------------------
# FUNÇÃO MENU PRINCIPAL
#-------------------------------------------------------------------------------------------------------------------------
def menu_principal():
    os.system('cls')
    print("                                                                                  ")
    print("------------------------------------------------------------------------------")
    print("                 ****   Menu de Investimentos  ****                               ", color='blue')
    print("                           Menu Principal                                         ", color='yellow')
    print("                      ------------------------                                    ")
    print("                     ")
    print("- Verificar Parâmetros - Filtro/Recomendação - Ações    ", tag=' 1 ', tag_color='green', color='cyan')
    print("- Restaurar Parâmetros - Filtro/Recomendação - Original ", tag=' 2 ', tag_color='green', color='cyan')
    print("- Alterar   Parâmetros - Mínimos e Máximos   - Ações    ", tag=' 3 ', tag_color='green', color='cyan')
    print("- Gerar Planilha Análise                     - FII      ", tag=' 4 ', tag_color='green', color='cyan')
    print("- Gerar Planilha Análise                     - Ações    ", tag=' 5 ', tag_color='green', color='cyan') 
    print("- Gerar Planilha Análise                     - Ações e FIIs", tag=' 6 ', tag_color='green', color='cyan')     
    print(" ")
    print("- Processar Dados(BD)                        - Ações    ", tag=' 7 ', tag_color='green', color='cyan')
    print("- Listar Dados   (BD)                        - CSV-File ", tag=' 8 ', tag_color='green', color='cyan')
    print("- Encerrar Programa                          - Sair     ", tag=' 9 ', tag_color='green', color='cyan')
    print("                     ")
    print("------------------------------------------------------------------------------")
    print("                ", menu_msg, color='purple')
    print("------------------------------------------------------------------------------")
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNÇÃO PARA PROCESSAR O MENU PRINCIPAL
#-------------------------------------------------------------------------------------------------------------------------
def processa_menu():
    menu_principal()
    opcao = int(input("Digite Opção Desejada :"))
    while opcao != 9:
        if opcao == 1:
            consulta_json()
            lista_json(opcao)
        elif opcao == 2:
            cria_json()
        elif opcao == 3:
            consulta_json()
            altera_min_max()
        elif opcao == 4:
            analise_fii()
        elif opcao == 5:
            analise_acoes()
        elif opcao == 6:
            analise_acoes_fii()
        elif opcao == 7:
            consulta_json()
            pesquisa_set()
        elif opcao == 8:
            exporta_csv()  
        else:
            print("                             ")
            print("Opção Indisponível no Momento", color='red')
            print("")
        menu_principal()
        opcao = int(input("Digite Opção Desejada : "))
    print("Encerrando o Programa - Até Logo", color='yellow')
    exit()
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA ALTERAR OS PARAMETROS MINIMOS E MAXIMOS UTILIZADOS COMO FILTRO DE SELECAO DAS ACOES
#-------------------------------------------------------------------------------------------------------------------------
def altera_min_max():
    opcao_parm = '0'
    cdy = 'cyan'
    cpl = 'cyan'
    cpvp = 'cyan'
    croe = 'cyan'
    cc5y = 'cyan'
    cliq2m = 'cyan'
    while opcao_parm != '9': 
        os.system('cls')

        print('----------------------------------------------------------------------------------')
        print('                     Valores Minimos e Máximos Utilizados                       ', color='yellow')
        print('         ')
        print(' - Valores Dividend Yield: ',  'Valor Mínimo =', xpap_dy_min, '     Valor Máximo =', xpap_dy_max, color=cdy, tag=1, tag_color='green')
        print(' - Valores P/L           : ',  'Valor Mínimo =', xpap_pl_min, '     Valor Máximo =', xpap_pl_max, color=cpl, tag=2, tag_color='green')
        print(' - Valores P/VP          : ',  'Valor Mínimo =', xpap_pvp_min, '     Valor Máximo =', xpap_pvp_max, color=cpvp, tag=3, tag_color='green') 
        print(' - Valores ROE           : ',  'Valor Mínimo =', xpap_roe_min, '    Valor Máximo =', xpap_roe_max, color=croe, tag=4, tag_color='green')
        print(' - Valor Crescimento Reco: ',  'Valor Mínimo =', xpap_c5y_min, color=cc5y, tag=5, tag_color='green')
        print(' - Valor Liquidez 2 meses: ',  'Valor Mínimo =', xpap_liq2m_min, color=cliq2m, tag=6, tag_color='green')
        print('         ')
        print(' - Voltar ao menu Principal ', color='cyan', tag=9, tag_color='yellow')
        print('         ')
        print('----------------------------------------------------------------------------------')
        opcao_parm = (input("Selecione o parâmetro a ser alterado: ")) 
        if opcao_parm == '1':
            opcao_parm_min = (input("Digite o valor minimo para DY : "))
            jsonf = 'jpap_dy_min'
            jsonv = float(opcao_parm_min)
            atualiza_json(jsonf, jsonv)
            opcao_parm_max = (input("Digite o valor máximo para DY : ")) 
            jsonf = 'jpap_dy_max'
            jsonv = float(opcao_parm_max)
            atualiza_json(jsonf, jsonv)
            cdy = 'green'
        consulta_json()
#
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA GERAR PLANILHA COM ANALISE DE ACOES E FII SEPARADAS POR ABAS
#-------------------------------------------------------------------------------------------------------------------------
def analise_acoes_fii():
    global menu_msg
    menu_msg = "   Opção Indisponível no Momento"
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA GERAR/EXPORTAR DADOS DA BASE DE DADOS UTILIZADOS NA OPCAO 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
def exporta_csv():
    con = sqlite3.connect("investdb.db", timeout=10)
    cursor = con.cursor()
    cursor.execute("Select * from tbinvest")
    with open("analise.csv", "w") as csv_file:
        csv_writer = csv.writer(csv_file, delimiter=";", lineterminator='\n')
        csv_writer.writerow([i[0] for i in cursor.description])
        csv_writer.writerows(cursor)
    print("Planilha Gerada com Sucesso !!!!")
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA PESQUISAR SETOR (BASE FUNDAMENTUS) PARA INICIAR ANALISE DE ACOES E CARREGAR BASE UTILIZADA NA OPC 7-EM DEUSO
#-------------------------------------------------------------------------------------------------------------------------
def pesquisa_set():
    s = 1
    while s < 99:
        tes = fundamentus.list_papel_setor(s)
        tam = len(tes)
 #       print('set=', s)
 #       print( 'tam=', tam)
        for x in tes:
#            print(x)
            papel = x
            pesquisa_dados(s, papel)
        s = s + 1
    opcao = 7
    lista_json(opcao)
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA PESQUISAR DADOS DAS ACOES DO SETOR PARA ANALISE DE ACOES E CARREGAR BASE - UTILIZADO NA OPC 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
def pesquisa_dados(s, papel):
#    print('rotina nova xxxx=', papel, s)
    con = sqlite3.connect("investdb.db", timeout=10)
    data = con.execute("select * from TBSET where cod_set = ? """, (s,))
    for row in data:
        print ("Setor=", row[0])
        print ("DESC=", row[1])
        print ("Complemento=", row[2], "\n")
        df4 = fundamentus.get_papel(papel)
        xcod_pap = papel
        xcod_set = s
        xdesc_set_res = row[1]
        xpap_cotacao = (df4.iloc[0].Cotacao)

#     TRATAMENTO PL        
        xpap_pl      = (df4.iloc[0].PL)
        try:
            xpap_pl      = int(xpap_pl)
            xpap_pl      = xpap_pl / 100
        except:
            xpap_pl      = 0

#     TRATAMENTO PVP
        xpap_pvp     = (df4.iloc[0].PVP)
        try:
            xpap_pvp     = int(xpap_pvp)
            xpap_pvp     = xpap_pvp / 100
        except:
            xpap_pvp     = 0

        xpap_psr	 = (df4.iloc[0].PSR)
        xpap_dy	     = (df4.iloc[0].Div_Yield)

#     TRATAMENTO PA    
        xpap_pa	     = (df4.iloc[0].PAtivos)
        try:
            xpap_pa      = int(xpap_pa)
            xpap_pa      = xpap_pa/100
        except:
            xpap_pa      = 0

#     TRATAMENTO  PCG
        xpap_pcg     = (df4.iloc[0].PCap_Giro)
        try:
            xpap_pcg     = int(xpap_pcg)
            xpap_pcg     = xpap_pcg / 100
        except:
            xpap_pcg     = 0

#     TRATAMENTO  PEBIT
        xpap_pebit	  = (df4.iloc[0].PEBIT)
        try:
            xpap_pebit = int(xpap_pebit)
            xpap_pebit = xpap_pebit / 100
        except:
            xpap_pebit = 0

#     TRATAMENTO  PACL
        xpap_pacl	  = (df4.iloc[0].PAtiv_Circ_Liq)
        try:
            xpap_pacl = int(xpap_pacl)
            xpap_pacl = xpap_pacl / 100
        except:
            xpap_pacl = 0

#     TRATAMENTO  EVEBIT
        xpap_evebit	  = (df4.iloc[0].EV_EBIT)
        try:
            xpap_evebit = int(xpap_evebit)
            xpap_evebit = xpap_evebit / 100
        except:
            xpap_evebit = 0

#     TRATAMENTO  EVEBITDA
        xpap_evebitda = (df4.iloc[0].EV_EBITDA)
        try:
            xpap_evebitda = int(xpap_evebitda)
            xpap_evebitda = xpap_evebitda / 100 
        except:
            xpap_evebitda = 0
     
        xpap_mrgebit  = (df4.iloc[0].Marg_EBIT)
        xpap_mrgliq   = (df4.iloc[0].Marg_Liquida)
        xpap_roic     = (df4.iloc[0].ROIC)
        xpap_roe      = (df4.iloc[0].ROE)

#     TRATAMENTO  LIQC
        xpap_liqc     = (df4.iloc[0].Liquidez_Corr)
        try:
            xpap_liqc = int(xpap_liqc)
            xpap_liqc = xpap_liqc / 100
        except:
            xpap_liqc = 0

#     TRATAMENTO  divbpatr
        xpap_divbpatr = (df4.iloc[0].Div_Br_Patrim)
        try:
            xpap_divbpatr = int(xpap_divbpatr)
            xpap_divbpatr = xpap_divbpatr / 100
        except:
            xpap_divbpatr = 0

 ## NAO ENCONTRADO CORRESPONDENTE PARA LIQ2M
 ## EH O CAMPO Vol_med_2m
        xpap_liq2m    = (df4.iloc[0].Vol_med_2m)

        xpap_patrliq   = (df4.iloc[0].Patrim_Liq)
        xpap_c5y      =  (df4.iloc[0].Cres_Rec_5a)
        xpap_rec      = "S"

## Timestamp
        xpap_timestamp = datetime.datetime.now()

        print("Papel=", xcod_pap)
        print("Setor=", xcod_set)
        print("Descr=", xdesc_set_res)
        print("Cotacao=", xpap_cotacao)
        print("PL=", xpap_pl)
        print("PVP=", xpap_pvp)
        print("DY=", xpap_dy)
        print("PA=", xpap_pa)
        print("PCG=", xpap_pcg)
        print("PEBIT=", xpap_pebit)
        print("PACL=", xpap_pacl)
        print("EVEBIT =", xpap_evebit)
        print("EVEBITDA =", xpap_evebitda)
        print("MRGEBIT =", xpap_mrgebit)
        print("MRGLIQ =", xpap_mrgliq)
        print("ROIC=", xpap_roic)
        print("ROE=", xpap_roe)
        print("liqc=", xpap_liqc)
        print("liq2m=", xpap_liq2m)
        print("patrliq=", xpap_patrliq)
        print("divbpatr=", xpap_divbpatr)
        print("c5y=", xpap_c5y)
        print("TS =", xpap_timestamp)

    # Tratamento do campo de Dividend Yield retirando o %
        if xcod_pap == "BBAS3":
            print("BBAS3 p1: ", xpap_rec)

        xpap_dy = xpap_dy.replace('%', ' ')
        xpap_dy = float(xpap_dy)
        if xpap_dy < xpap_dy_min or xpap_dy > xpap_dy_max:
            xpap_rec      = "N"

        if xcod_pap == "BBAS3":
            print("BBAS3 p2: ", xpap_rec)

        if xpap_pl < xpap_pl_min or xpap_pl > xpap_pl_max:
            xpap_rec      = "N"

        if xcod_pap == "BBAS3":
            print("BBAS3 p3: ", xpap_rec)

        if xpap_pvp < xpap_pvp_min or xpap_pvp > xpap_pvp_max:
            xpap_rec      = "N"
            print(xpap_pvp, xpap_pvp_min, xpap_pvp_max)

        if xcod_pap == "BBAS3":
            print("BBAS3 p4: ", xpap_rec)
        xpap_roe = xpap_roe.replace('%', ' ')
        try:
            xpap_roe = float(xpap_roe)
        except:
            xpap_roe = float(0.0)

        if xpap_roe < xpap_roe_min or xpap_roe > xpap_roe_max:
            xpap_rec      = "N"

        if xcod_pap == "BBAS3":
            print("BBAS3 p5: ", xpap_rec)

        xpap_liq2m = int(xpap_liq2m)
        if xpap_liq2m < xpap_liq2m_min:
            xpap_rec      = "N"

        if xcod_pap == "BBAS3":
            print("BBAS3 p6: ", xpap_rec)
            xTEST = input("VERIFICAR MENSAGEM UPD - PRESSIONAR ENTER: ")
        xpap_c5y = xpap_c5y.replace('%', ' ')
        try:
            xpap_c5y = float(xpap_c5y)
        except:
            xpap_c5y = float(0.0)   

        if xpap_c5y < xpap_c5y_min:
            xpap_rec      = "N"
# Inclusão de dados Papel      
        try:
            ipapeldados = (xcod_pap, xcod_set)
#           conn = sqlite3.connect("investdb.db", timeout=10)
            insert_papel(con, ipapeldados)
        except Exception as e: 
            print(e)
        print ('-----------------------------------------')
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA CARREGAR DADOS NA BASE DE DADOS SQLITE - UTILIZADO NA  OPCAO 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
# Atualização de dados Papel
        try:
            papeldados = (xcod_set, xdesc_set_res, xpap_cotacao, xpap_pl, xpap_pvp, xpap_psr, xpap_dy, xpap_pa,
                          xpap_pcg, xpap_pebit, xpap_pacl, xpap_evebit, xpap_evebitda, xpap_mrgebit, xpap_mrgliq, 
                          xpap_roic, xpap_roe, xpap_liqc, xpap_liq2m, xpap_patrliq, xpap_divbpatr, xpap_c5y, xpap_rec,
                          xpap_timestamp, xcod_pap)
#            conn = sqlite3.connect("investdb.db", timeout=10)
            update_papel(con, papeldados)
        except Exception as e: print(e)
        print ('-----------------------------------------')

#        xTEST = input("VERIFICAR MENSAGEM UPD - PRESSIONAR ENTER: ")
    con.close()     
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA GERAR BASE DE DADOS SQLITE UTILIZADOS NA OPCAO 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
def criar_tab():
    try:
        print("inicio rotina criar tabela")
        con = sqlite3.connect("investdb.db")
        cur = con.cursor()
        cur.execute("CREATE TABLE setores(codigo, setdes1, setdes2)")
        res = cur.execute("Select name from sqlite_master")
        mav = res.fetchone()
        print(mav)
    except:
        print("Erro criação tabela setores")
#
def exibe_papeis():
    df = fundamentus.get_resultado()
    print("****************************************************************")   
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO DESATIVADA PARA PESQUISAR DADOS DE UMA ACAO ESPECIFICA - FUNDAMENTUS
#-------------------------------------------------------------------------------------------------------------------------
def pesquisa_papel():
    acao = input("Qual papel pesquisar: ")
    df2 = fundamentus.get_papel(acao)
# para usar iloc importar pandas  ... iloc[0] irá listar primeira linha do dataframe
#   first_row = df3.iloc[0]
#   print(first_row)
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA ATUALIZAR BASE DE SETOR SQLITE - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
def update_setor(conn, setdados):
    sql_update = ''' UPDATE tbset 
    SET desc_set_res = ?
    WHERE cod_set = ?
    '''

    try:
        cursor = conn.cursor()
        cursor.execute(sql_update, setdados)
        qtd = cursor.rowcount
        conn.commit()
 #      conn.close() 
    except sqlite3.Error:
        print("Erro Update")
#    except Exception as e:
#        print(e)   
#except sqlite3.Error as er:
#        print("Failed to update sqlite table")
    finally: 
        print('Execute Update SQL successfully - ALTERADOS=.', qtd)
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA ATUALIZAR DADOS DA BASE DE DADOS UTILIZADOS NA OPCAO 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
def update_papel(conn, papeldados):
    sql_update = ''' UPDATE tbinvest
    SET cod_set = ?,
        desc_set_res = ?,
        pap_cotacao = ?,
        pap_pl = ?,
        pap_pvp = ?,
        pap_psr = ?,
        pap_dy = ?,
        pap_pa = ?,
        pap_pcg = ?,
        pap_pebit = ?,
        pap_pacl = ?,
        pap_evebit = ?,
        pap_evebitda = ?,
        pap_mrgebit = ?,
        pap_mrgliq = ?,
        pap_roic = ?,
        pap_roe = ?,
        pap_liqc = ?,
        pap_liq2m = ?,
        pap_patrliq = ?,
        pap_divbpatr = ?,
        pap_c5y	= ?,
        pap_rec = ?,
        pap_ts = ?
    WHERE cod_pap = ?
    '''

    try:
        cursor = conn.cursor()
        cursor.execute(sql_update, papeldados)
        qtdx = cursor.rowcount
        conn.commit()
#       conn.close() 
#    except sqlite3.Error:
#        print("Erro Update tbinvest")
    except Exception as e:
         print(e)   
#except sqlite3.Error as er:
#        print("Failed to update sqlite table ")
    finally: 
        print('Execute Update tbinvest SQL successfully - ALTERADOS=.', qtdx)
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO INSETIR DADOS DA TABELA DE SETOR/ACAO DA BASE DE DADOS UTILIZADOS NA OPCAO 7 - EM DESUSO
#-------------------------------------------------------------------------------------------------------------------------
##Inicio Rotina de Inclusão do Papel na tabela TBINVEST
def insert_papel(connn, ipapeldados):
    sql_insert =  ''' INSERT into tbinvest (cod_pap, cod_set) values (?,?) ''' 
    try:
        cursor = connn.cursor()
        cursor.execute(sql_insert, ipapeldados)
        qtdz = cursor.rowcount
        connn.commit()
#    except sqlite3.Error:
#        print("Erro Insert tbinvest")
    except Exception as e:
         print(e)   
#except sqlite3.Error as er:
#        print("Failed to Insert sqlite table ")
    finally: 
        print('Execute Insert tbinvest SQL successfully - Incluídos=.', qtdz)
##Fim Rotina de Inclusão do Papel na tabela TBINVEST
 #
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO UTILIZADA PARA DAR CARGA INICIAL NA BASE JSON COM PARAMETROS DE FILTRO DE SELECAO DAS ACOES
#-------------------------------------------------------------------------------------------------------------------------       
def cria_json():
    dados = {
    "jpap_dy_fil":'S',
    "jpap_dy_min":7.0,
    "jpap_dy_max":14.0,
    "jpap_pl_fil":'S',
    "jpap_pl_min":3.0,
    "jpap_pl_max":10.0,
    "jpap_pvp_fil":'S',	
    "jpap_pvp_min":0.5,
    "jpap_pvp_max":2.0,	
    "jpap_roe_fil":'S',
    "jpap_roe_min":15.0,
    "jpap_roe_max":30.0,
    "jpap_liq2m_fil":'S',
    "jpap_liq2m_min":1000000,	
    "jpap_c5y_fil":'S',			 
    "jpap_c5y_min":10.0
}
    arquivo = open("dados.json", "w")
    json.dump(dados, arquivo)
    arquivo.close()
    global menu_msg
    menu_msg = "Parâmetros Originais Restaurados"
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA ATUALIZAR BASE JSON COM PARAMETROS DE FILTRO DE SELECAO DAS ACOES
#-------------------------------------------------------------------------------------------------------------------------
def atualiza_json(jsonf, jsonv):
    with open('dados.json', 'r') as f:
        json_data = json.load(f)
        json_data[jsonf] = jsonv
    with open('dados.json', 'w') as f:
        f.write(json.dumps(json_data))
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA CONSULTAR BASE JSON COM PARAMETROS DE FILTRO DE SELECAO DAS ACOES
#-------------------------------------------------------------------------------------------------------------------------
def consulta_json():
    arquivo = open("dados.json", "r")
    dados = json.load(arquivo)
    arquivo.close()
    print(dados)
    print('teste teste teste parm')
    print(dados["jpap_dy_max"])
    global xpap_dy_fil
    xpap_dy_fil = (dados["jpap_dy_fil"])
    global xpap_dy_min
    xpap_dy_min = (dados["jpap_dy_min"])
    global xpap_dy_max
    xpap_dy_max = (dados["jpap_dy_max"])
    global xpap_pl_fil
    xpap_pl_fil = (dados["jpap_pl_fil"])
    global xpap_pl_min
    xpap_pl_min = (dados["jpap_pl_min"])
    global xpap_pl_max
    xpap_pl_max = (dados["jpap_pl_max"])
    global xpap_pvp_fil
    xpap_pvp_fil = (dados["jpap_pvp_fil"])
    global xpap_pvp_min
    xpap_pvp_min = (dados["jpap_pvp_min"])
    global xpap_pvp_max
    xpap_pvp_max = (dados["jpap_pvp_max"])
    global xpap_roe_fil
    xpap_roe_fil = (dados["jpap_roe_fil"])
    global xpap_roe_min
    xpap_roe_min = (dados["jpap_roe_min"])
    global xpap_roe_max
    xpap_roe_max = (dados["jpap_roe_max"])
    global xpap_liq2m_fil
    xpap_liq2m_fil = (dados["jpap_liq2m_fil"])
    global xpap_liq2m_min
    xpap_liq2m_min = (dados["jpap_liq2m_min"])		 
    global xpap_c5y_fil
    xpap_c5y_fil = (dados["jpap_c5y_fil"])
    global xpap_c5y_min
    xpap_c5y_min = (dados["jpap_c5y_min"])
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA LISTAR BASE JSON COM PARAMETROS DE FILTRO DE SELECAO DAS ACOES
#-------------------------------------------------------------------------------------------------------------------------
def lista_json(opcao):
    if xpap_dy_fil == 's' or xpap_dy_fil == 'S':
        dy_fil = '[Habilitado]'
        dy_filc = 'green'
    else:
        dy_fil = '[Desabilitado]'
        dy_filc = 'red'
    if xpap_pl_fil == 's' or xpap_pl_fil == 'S':
        pl_fil = '[Habilitado]'
        pl_filc = 'green'
    else:
        pl_fil = '[Desabilitado]'
        pl_filc = 'red'
    if xpap_pvp_fil == 's' or xpap_pvp_fil == 'S':
        pvp_fil = '[Habilitado]'
        pvp_filc = 'green'
    else:
        pvp_fil = '[Desabilitado]'
        pvp_filc = 'red'
    if xpap_roe_fil == 's' or xpap_roe_fil == 'S':
        roe_fil = '[Habilitado]'
        roe_filc = 'green'
    else:
        roe_fil = '[Desabilitado]'
        roe_filc = 'red'
    if xpap_c5y_fil == 's' or xpap_c5y_fil == 'S':
        c5y_fil = '[Habilitado]'
        c5y_filc = 'green'
    else:
        c5y_fil = '[Desabilitado]'
        c5y_filc = 'red'
    if xpap_liq2m_fil == 's' or xpap_liq2m_fil == 'S':
        liq2m_fil = '[Habilitado]'
        liq2m_filc = 'green'
    else:
        liq2m_fil = '[Desabilitado]'
        liq2m_filc = 'red'
    
    print('----------------------------------------------------------------------------------')
    print('                            Parâmetros Utilizados                                 ', color='yellow')
    print('         ')
    print('Filtro Dividend Yield    :' ,xpap_dy_fil, '      ', dy_fil, color=dy_filc)
    print('Dividend Yield Mínimo    :' ,xpap_dy_min)
    print('Dividend Yield Máximo    :' ,xpap_dy_max)
    print('         ')
    print('Filtro P/L               :' ,xpap_pl_fil, '      ', pl_fil, color=pl_filc)
 
    print('           PL  Mínimo    :' ,xpap_pl_min)
    print('           PL  Máximo    :' ,xpap_pl_max)
    print('         ')
    print('Filtro P/VP              :' ,xpap_pvp_fil, '      ', pvp_fil, color=pvp_filc)
    print('          PVP  Mínimo    :' ,xpap_pvp_min)
    print('          PVP  Máximo    :' ,xpap_pvp_max)
    print('         ')
    print('Filtro ROE               :' ,xpap_roe_fil, '      ', roe_fil, color=roe_filc)
    print('          ROE  Mínimo    :' ,xpap_roe_min)
    print('          ROE  Máximo    :' ,xpap_roe_max)
    print('         ')
    print('Filtro Crescimento Recor :' ,xpap_c5y_fil, '      ', c5y_fil, color=c5y_filc)
    print('Cresc. Rec. Mínimo       :' ,xpap_c5y_min)
    print('         ')
    print('Filtro Liquidez 2 meses  :' ,xpap_liq2m_fil, '      ', liq2m_fil, color=liq2m_filc)
    print('Liq2m Mínimo             :' ,xpap_liq2m_min)
    print('         ')
    print('----------------------------------------------------------------------------------')
    if opcao == 1:
        opcao_parm = (input("Deseja  Alterar(Habilitar/Desabilitar)  Filtros   [S/N]? "))
        dy_flag = ''
        pl_flag = ''
        pvp_flag = ''
        roe_flag = ''
        c5y_flag = ''
        liq2m_flag =''
        if opcao_parm == 's' or opcao_parm == "S":
            dy_flag = (input("Deseja    Habilitar    filtro   Dividend    Yield [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if dy_flag == 's' or dy_flag == 'S' or dy_flag == 'n' or dy_flag == 'N':
            jsonf = 'jpap_dy_fil'
            jsonv = dy_flag
            atualiza_json(jsonf, jsonv)
#
#       verificando se altera filtro P/L
        if opcao_parm == 's' or opcao_parm == "S":
            pl_flag = (input("Deseja    Habilitar    filtro               P / L [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if pl_flag == 's' or pl_flag == 'S' or pl_flag == 'n' or pl_flag == 'N':
            jsonf = 'jpap_pl_fil'
            jsonv = pl_flag
            atualiza_json(jsonf, jsonv)    
#
#       verificando se altera filtro P/VP
        if opcao_parm == 's' or opcao_parm == "S":
            pvp_flag = (input("Deseja   Habilitar    filtro               P / VP [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if pvp_flag == 's' or pvp_flag == 'S' or pvp_flag == 'n' or pvp_flag == 'N':
            jsonf = 'jpap_pvp_fil'
            jsonv = pvp_flag
            atualiza_json(jsonf, jsonv)    
#
#       verificando se altera filtro ROE
        if opcao_parm == 's' or opcao_parm == "S":
            roe_flag = (input("Deseja   Habilitar   filtro                   ROE [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if roe_flag == 's' or roe_flag == 'S' or roe_flag == 'n' or roe_flag == 'N':
            jsonf = 'jpap_roe_fil'
            jsonv = roe_flag
            atualiza_json(jsonf, jsonv)
#
#       verificando se altera filtro c5y
        if opcao_parm == 's' or opcao_parm == "S":
            c5y_flag = (input("Deseja  Habilitar  filtro Crescimento  Recorrente [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if c5y_flag == 's' or c5y_flag == 'S' or c5y_flag == 'n' or c5y_flag == 'N':
            jsonf = 'jpap_c5y_fil'
            jsonv = c5y_flag
            atualiza_json(jsonf, jsonv)
#
#       verificando se altera filtro liq2m
        if opcao_parm == 's' or opcao_parm == "S":
            liq2m_flag = (input("Deseja  Habilitar filtro Liquidez Últimos 2 Meses [S/N]? "))
            print('----------------------------------------------------------------------------------') 
        if liq2m_flag == 's' or liq2m_flag == 'S' or liq2m_flag == 'n' or liq2m_flag == 'N':
            jsonf = 'jpap_liq2m_fil'
            jsonv = liq2m_flag
            atualiza_json(jsonf, jsonv)
    else:
        opcao = input("Pressione qualquer tecla para voltar ao menu")
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA CONSULTAR E SELECIONAR FUNDOS DE INVESTIMENTO IMOBILIARIOS - OPCAO 4
#-------------------------------------------------------------------------------------------------------------------------
def analise_fii():
    url = "https://www.fundamentus.com.br/fii_resultado.php"
    headers = {'User-Agent':'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'}
    r = requests.get(url, headers=headers)
    tabela = pd.read_html(StringIO(r.text))
 
 ## tabela = pd.read_html(r.text)
    df5 = tabela[0]
    print('tam=', df5.size, df5.info)
    print('len=', len(df5))
    contador = len(df5)
    # substituir espaçocs por _
#   df5.columns = [c.replace(' ', '_') for c in df5.columns]
    df5 = df5.rename(columns={'FFO Yield': 'FFO_Yield'})
    df5 = df5.rename(columns={'Dividend Yield': 'dividend_yield'}) 
    df5 = df5.rename(columns={'P/VP': 'p_vp'})
    df5 = df5.rename(columns={'Valor de Mercado': 'valor_mercado'})
    df5 = df5.rename(columns={'Liquidez': 'liquidez'})
    df5 = df5.rename(columns={'Qtd de imóveis':  'qtd_imoveis'})
    df5 = df5.rename(columns={'Preço do m2': 'preco_m2'})
    df5 = df5.rename(columns={'Aluguel por m2': 'aluguel_m2'})
    df5 = df5.rename(columns={'Cap Rate': 'cap_rate'})
    df5 = df5.rename(columns={'Vacância Média': 'vacancia_media'})
    m = 0
    print('colunas=', df5.columns)
    while(m < contador):
#     TRATAMENTO Cotação        
        xfii_cot         = (df5.iloc[m].Cotação)
        try:
            xfii_cot      = int(xfii_cot)
            xfii_cot      = xfii_cot / 100
        except:
            xfii_cot      = 0
#     TRATAMENTO FFO Yield
        xfii_ffo         = (df5.iloc[m].FFO_Yield)
        xfii_ffo = xfii_ffo.replace('%', ' ')
        xfii_ffo = xfii_ffo.replace(',', '.')
        print('xxxx=', xfii_ffo)
        try:
            xfii_ffo = float(xfii_ffo)
        except:
            xfii_ffo = float(0.0)
#        xfii_dy = (df5.iloc[m].dividend_yield)
        print('papel=', df5.loc[m].Papel, 'cotacao=', xfii_cot, 'FFO Yield=', xfii_ffo)
        print('dy   =', df5.loc[m].dividend_yield, 'P/VP=', df5.loc[m].p_vp, 'Valor de Mercado=', df5.loc[m].valor_mercado)
        print('Liquidez  =', df5.loc[m].liquidez, 'Qtd de imóveis=', df5.loc[m].qtd_imoveis, 'Preço do m2=', df5.loc[m].preco_m2)
        print('Aluguel por m2  =', df5.loc[m].aluguel_m2, 'Cap Rate=', df5.loc[m].cap_rate, 'Vacancia Média=', df5.loc[m].vacancia_media)
        print('----------------------------------------------') 
        m = m + 1
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO PARA CONSULTAR E SELECIONAR ACOES - OPCAO 5
#-------------------------------------------------------------------------------------------------------------------------
def analise_acoes():
# definindo variárives do workbook, criando planilha e gravando cabeçalho
    wb = openpyxl.Workbook()
    wb.create_sheet('acoes')
    wb.remove_sheet(wb['Sheet'])
    planilha_acoes = wb['acoes']
    plan_header = ['Papel', 'Cotação',	'P/L', 'P/VP', 'PSR', 'Div.Yield', 'P/Ativo', 'P/Cap.Giro',	'P/EBIT', 'P/Ativ Circ.Liq', 'EV/EBIT',	'EV/EBITDA', 'Mrg Ebit', 'Mrg. Líq.', 'Liq. Corr.', 'ROIC', 'ROE', 'Liq.2meses',	'Patrim. Líq', 'Dív.Brut/ Patrim.',	'Cresc. Rec.5a', 'flag']
    planilha_acoes.append(plan_header)
# executar rotina para obter parametros de análise para definir flag de recomendação
    consulta_json()    
# definindo origem da leitura de dados, site fundamentus ações
    url = "https://www.fundamentus.com.br/resultado.php"
    headers = {'User-Agent':'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'}
    r = requests.get(url, headers=headers)
    tabela = pd.read_html(StringIO(r.text))
 ## tabela = pd.read_html(r.text)
    df6 = tabela[0]
    print('tam=', df6.size, df6.info)
    print('len=', len(df6))
    contador = len(df6)
    # substituir espaçocs por _
# formatando colunas da tabela dataframe
    df6 = df6.rename(columns={'Papel': 'a_papel'})
    df6 = df6.rename(columns={'Cotação': 'a_cotacao'})	
    df6 = df6.rename(columns={'P/L': 'a_pl'})	
    df6 = df6.rename(columns={'P/VP': 'a_pvp'})	
    df6 = df6.rename(columns={'PSR': 'a_psr'})
    df6 = df6.rename(columns={'Div.Yield': 'a_divyield'})
    df6 = df6.rename(columns={'P/Ativo': 'a_pativo'})
    df6 = df6.rename(columns={'P/Cap.Giro': 'a_capgiro'})	
    df6 = df6.rename(columns={'P/EBIT': 'a_pebit'})
    df6 = df6.rename(columns={'P/Ativ Circ.Liq': 'a_pativcircliq'})
    df6 = df6.rename(columns={'EV/EBIT': 'a_evebit'})
    df6 = df6.rename(columns={'EV/EBITDA': 'a_evebitda'})
    df6 = df6.rename(columns={'Mrg Ebit': 'a_mrgebit'})
    df6 = df6.rename(columns={'Mrg. Líq.': 'a_mrgliq'})
    df6 = df6.rename(columns={'Liq. Corr.': 'a_liqcorr'})
    df6 = df6.rename(columns={'ROIC': 'a_roic'})
    df6 = df6.rename(columns={'ROE': 'a_roe'})
    df6 = df6.rename(columns={'Liq.2meses': 'a_liq2meses'})
    df6 = df6.rename(columns={'Patrim. Líq': 'a_patrimliq'})
    df6 = df6.rename(columns={'Dív.Brut/ Patrim.': 'a_divbrutpatrim'})
    df6 = df6.rename(columns={'Cresc. Rec.5a': 'a_crescrec5a'})
    m = 0
    print('colunas=', df6.columns)
    while(m < contador):
        a_papel            = (df6.iloc[m].a_papel) 
#     TRATAMENTO Cotação        
        a_cotacao          = (df6.iloc[m].a_cotacao) 
        try:
            a_cotacao    = int(a_cotacao)
            a_cotacao    = a_cotacao / 100
        except:
            a_cotacao      = 0
#     tratamento PL
        a_pl          = (df6.iloc[m].a_pl)
        a_pl          = a_pl.replace('.','')
        a_pl          = a_pl.replace(',','')
#       a_pl          = a_pl.replace('-','')
        a_pl          = int(a_pl)
        try:
             a_pl      = a_pl / 100
        except:
             a_pl      = 0
#tratamento PVP
        a_pvp            = (df6.iloc[m].a_pvp) 
        try:
            a_pvp      = int(a_pvp)
            a_pvp      = a_pvp / 100
        except:
            a_pvp      = 0
#tratamento PSR
        a_psr          = (df6.iloc[m].a_psr)
        try:
            a_psr      = int(a_psr)
            a_psr      = a_psr / 1000
        except:
            a_psr      = 0

#tratamento p ativos
        a_pativo           = (df6.iloc[m].a_pativo) 
        try:
            a_pativo    = int(a_pativo)
            a_pativo    = a_pativo / 1000
        except:
            a_pativo    = 0
#tratamento capital de giro
        a_capgiro      = (df6.iloc[m].a_capgiro) 
        try:
            a_capgiro  = int(a_capgiro)
            a_capgiro  = a_capgiro / 100
        except:
            a_capgiro  = 0
#tratamento p EBIT
        a_pebit            = (df6.iloc[m].a_pebit) 
        try:
            a_pebit  = int(a_pebit)
            a_pebit  = a_pebit / 100
        except:
            a_pebit  = 0
# TRATAMENTO  P ATIV CIRC LIQ
        a_pativcircliq     = (df6.iloc[m].a_pativcircliq) 		
        try:
            a_pativcircliq = int(a_pativcircliq)
            a_pativcircliq = a_pativcircliq / 100
        except:
             a_pativcircliq = 0
# TRATAMENTO  EV/EBIT		
        a_evebit           = (df6.iloc[m].a_evebit) 
        try:
            a_evebit = int(a_evebit)
            a_evebit = a_evebit / 100
        except:
        	a_evebit = 0
# TRATAMENTO  EV/EBITDA	
        a_evebitda = (df6.iloc[m].a_evebitda)
        try:
            a_evebitda = int(a_evebitda)
            a_evebitda = a_evebitda / 100
        except:
        	a_evebitda = 0
# TRATAMENTO Liquido corrente	
        a_liqcorr      = (df6.iloc[m].a_liqcorr)
        try:
            a_liqcorr = float(a_liqcorr)
            a_liqcorr = a_liqcorr / 100
        except:
        	a_liqcorr = 0
# TRATAMENTO Liq 2 meses
        a_liq2meses   = (df6.iloc[m].a_liq2meses)
        try:
            a_liq2meses   = a_liq2meses.replace('.','')
            a_liq2meses   = a_liq2meses.replace(',','')
            a_liq2meses = int(a_liq2meses)
            a_liq2meses = a_liq2meses / 100
        except:
        	a_liq2meses = 0
# TRATAMENTO Patrimonio Liquido
        a_patrimliq    = (df6.iloc[m].a_patrimliq)
        try:
            a_patrimliq   = a_patrimliq.replace('.','')
            a_patrimliq   = a_patrimliq.replace(',','')
            a_patrimliq = int(a_patrimliq)
            a_patrimliq = a_patrimliq / 100
        except:
        	a_patrimliq = 0
# Tratamento div.bruta/Patrim
        a_divbrutpatrim     = (df6.iloc[m].a_divbrutpatrim)
        try:
            a_divbrutpatrim = int(a_divbrutpatrim)
            a_divbrutpatrim = a_divbrutpatrim / 100
        except:
        	a_divbrutpatrim = 0       
#tratamento dividend yield
        a_divyield = (df6.iloc[m].a_divyield)
        try:
            a_divyield = a_divyield.replace('%','')
            a_divyield = a_divyield.replace(',','')
            a_divyield = int(a_divyield)
            a_divyield = a_divyield / 100
        except:
            a_divyield = 0
#tratamento ROE
        a_roe = (df6.iloc[m].a_roe)
        try:
            a_roe = a_roe.replace('%','')
            a_roe = a_roe.replace(',','')
            a_roe = int(a_roe)
            a_roe = a_roe / 100
        except:
            a_roe = 0
#tratamento a_crescrec5a
        a_crescrec5a = (df6.iloc[m].a_crescrec5a)
        try:
            a_crescrec5a = a_crescrec5a.replace('%','')
            a_crescrec5a = a_crescrec5a.replace(',','')
            a_crescrec5a = int(a_crescrec5a)
            a_crescrec5a = a_crescrec5a / 100
        except:
            a_crescrec5a = 0

#sem alterações no primeiro momento
        a_mrgebit = (df6.iloc[m].a_mrgebit)
        a_mrgliq = (df6.iloc[m].a_mrgliq)
        a_roic = (df6.iloc[m].a_roic)
        print('a_papel        =', a_papel)        
        print('a_cotacao	  =', a_cotacao)	  
        print('a_pl           =', a_pl)           
        print('a_pvp	      =', a_pvp)	      
        print('a_psr          =', a_psr)
        print('a_divyield     =', a_divyield)     
        print('a_pativo       =', a_pativo)       
        print('a_capgiro      =', a_capgiro)      
        print('a_pebit        =', a_pebit)        
        print('a_pativcircliq =', a_pativcircliq) 
        print('a_evebit       =', a_evebit)       
        print('a_evebitda     =', a_evebitda)     
        print('a_mrgebit      =', a_mrgebit)       
        print('a_mrgliq       =', a_mrgliq)      
        print('a_liqcorr      =', a_liqcorr)      
        print('a_roic         =', a_roic)         
        print('a_roe          =', a_roe)          
        print('a_liq2meses    =', a_liq2meses)    
        print('a_patrimliq    =', a_patrimliq)   
        print('a_divbrutpatrim=', a_divbrutpatrim) 
        print('a_crescrec5a   =', a_crescrec5a)   
        print('---------------------------------------------------------------------------------')
# definindo valor final da variavel flag de recomedação de compra
        a_flag = "S"
        if xpap_dy_fil == 's' or xpap_dy_fil == 'S' and a_divyield < xpap_dy_min or a_divyield > xpap_dy_max:
            a_flag      = "N"
        if xpap_pl_fil == 's' or xpap_pl_fil == 'S' and a_pl < xpap_pl_min or a_pl > xpap_pl_max:
            a_flag = "N"
        if xpap_pvp_fil == 's' or xpap_pvp_fil == 'S' and a_pvp < xpap_pvp_min or a_pvp > xpap_pvp_max:
            a_flag      = "N"
        if xpap_roe_fil == 's' or xpap_roe_fil == 'S' and a_roe < xpap_roe_min or a_roe > xpap_roe_max:
            a_flag      = "N"
        if xpap_liq2m_fil == 's' or xpap_liq2m_fil == 'S' and a_liq2meses < xpap_liq2m_min:
            a_flag      = "N"
        if xpap_c5y_fil == 's' or xpap_c5y_fil == 'S' and a_crescrec5a < xpap_c5y_min:
            a_flag  = "N"
        dadosac =  ([a_papel, a_cotacao, a_pl, a_pvp, a_psr, a_divyield, a_pativo, a_capgiro, a_pebit, a_pativcircliq, a_evebit, a_evebitda, a_mrgebit, a_mrgliq, a_liqcorr, a_roic, a_roe, a_liq2meses, a_patrimliq, a_divbrutpatrim , a_crescrec5a, a_flag])
        m = m + 1
        grava_xls_acoes(dadosac, planilha_acoes)
    wb.save('teste.xlsx')
    global menu_msg
    menu_msg = 'Planilha com ações selecionadas Gerada'
#
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO GRAVAR DADOS DAS ACOES NA PLANILHA - OPCAO 5
#-------------------------------------------------------------------------------------------------------------------------
def grava_xls_acoes(dadosac, planilha_acoes):
    d = (dadosac)
    planilha_acoes.append(d)
## 
#-------------------------------------------------------------------------------------------------------------------------
# FUNCAO - PROCESSAMENTO INICIAL DO PROGRAMA
#-------------------------------------------------------------------------------------------------------------------------
global menu_msg
menu_msg = '          Seja Bem Vindo'
processa_menu()
#criar_tab()
#exibe_papeis()
#pesquisa_papel()
#pesquisa_set()
