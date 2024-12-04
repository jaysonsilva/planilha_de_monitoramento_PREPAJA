#!/usr/bin/env python
# coding: utf-8

# In[ ]:


def main():
    import subprocess 
    bibliotecas = ["pandas", "pyodbc", "sqlalchemy", "openpyxl", "xlwings"] 
    for biblioteca in bibliotecas:
        try:
            __import__(biblioteca)  
        except ImportError: 
            subprocess.call(['pip', 'install', biblioteca])
    

    import ctypes

    import xlwings as xw

    app = xw.apps.active

    msg = 'Recomeda-se que somente uma planilha\nresumo esteja aberta nesse momento.\n\nClique "Ok" caso queira prosseguir com a extração.'
    title = 'Importante!'

    resposta = ctypes.windll.user32.MessageBoxW(0,msg,title,1)
    if resposta == 1:
        import pyodbc # Acessa servido
        import pandas as pd # Pacote de manipulação de dados
        import sqlalchemy as sa # para usar funções sql
        import openpyxl
        from sqlalchemy import event
        from sqlalchemy.engine import URL
        from sqlalchemy import create_engine
        from openpyxl import Workbook
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl import load_workbook
        import xlwings as xw

        import os


        #Cria o caminho para do arquivo
        # Iterar sobre todos os livros abertos no Excel
        for endereco in xw.books:
        # Verificar se o nome do livro começa com "testeEndereco"
            if endereco.name.startswith("PREPAJA"):
            # Se o nome do livro começar com "planilha", atribua-o à variável planilha
                planilha = endereco
            # Encerra o loop, pois já encontramos o livro desejado
                break

        planilha.activate()
        # Obtém o caminho completo do arquivo Excel
        caminho_arquivo = planilha.fullname
        #caminho = planilha.fullname
        #tirar = "S:"
        #Caminho = caminho.replace(tirar, "", 1)

        #Pega os dados na plninha para realizar a pesquisa no servidor 
        dadoEmpresa= pd.read_excel(caminho_arquivo, sheet_name='CAPA',usecols=[2],skiprows=11, nrows=21,header=None)
        dadoData= pd.read_excel(caminho_arquivo, sheet_name='CAPA',usecols=[2],skiprows=7, nrows=2,header=None)

        empresaPD = dadoEmpresa[2].dropna().astype(int).tolist()
        ano = dadoData.iloc[0, 0]
        mes = dadoData.iloc[1, 0]


        #Conecta com o servidor
        connection_string = 'Driver={SQL Server};Server=SERVIDOR_NAME,5468;Database=DATABASE_NAME;UID=USER;PWD=PASSWORD;'
        connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
        engine = create_engine(connection_url)
        conn = pyodbc.connect(connection_string)

        #Cursor de navegação nas tabelas
        cursor = conn.cursor()

        #Busca os dadoas da empresa no servidor
        sSQL = """ 
        SELECT 
          [sigla]
          ,year ([dataRevrea]) as Ano
          ,[IdAgente]
          ,[idSREAg]
          ,[id_natureza_dado]
          ,[id_detalhe_natureza]
          ,[idDet1]
          ,[idDet2]
          ,[idDet3]
          ,[id_grupo_tarifa]
          ,[id_detalhe_grupo_tarifa]
          ,[id_tipo_tarifa]
          ,[id_subgrupo] 
          ,[id_posto]
          ,[id_unidade]
          ,[id_UC]
          ,[valor]                             
          ,[DatadeRegistro]
        FROM [SGT_DEV].[BADNET].[BADNETFilt]/*[TabelaBADNET]*/
        WHERE (year(dataRevrea)= ? and month(dataRevrea)= ? and idAgente = ? /*and versao=17*/) and
        ((id_Natureza_dado in (1,2,3,4,7) and 
        id_Detalhe_Natureza in (1,2,3,4,99) and 
        idDet1 in (120,51,52,50,22,20,999,205,6,4,8,17,15,196,203,12,24,130,14,70,71,24,50,205) and 
         idDet2 in (999999,73,82,55,74,81,53,54,50,51,100,101,60,62,106,52,350,104,150,176,250,247,22,102,63,61,79,143,140,144,30,29,82,86,52,350) and 
         idDet3 in (999,179,178,242,235,241,240,106,220,523,524,525,526,82,81,503,512,73,102,120,123,167,510,518,529,512,234,74,502,515,516,514,180,519,522,110,111,182,183,184,112,198,11,201,527,528,529,530,531,532,533)   and 
         id_Grupo_TARIFA IN (99,1,2,3,4) and 
         id_Detalhe_Grupo_TARIFA IN (999,14,15,28,29,25,17,18,1,2,3,7,11,9,8,10,24,16,19,12,30) and id_Tipo_Tarifa IN (1,2,3,9) and 
         id_SUBGRUPO IN (20,21,22,99,3,4,5,2,1,3,6,7,8,9,10) and 
         id_POSTO IN (9) and 
         id_Unidade IN (4,3,2,99,5) and 
         id_UC IN  (99999,391,5707,0,72,385,383,69,63,39,404,405,6587,44,382,47,40,43,370,5697,5368,6898,2866,396,390,5216,86,400) ))  
                """
        #Construção do dataframe dos dados das empresas
        resultados=[]
        for empresa in empresaPD:   
            cursor.execute(sSQL, (int(ano), int(mes), int(empresa) ))
            resultados_empresa = cursor.fetchall()
            resultados.append(resultados_empresa)


        #ajustes no dataframe 
        for lista_por_empresa in resultados:
            for lista_de_id in lista_por_empresa:
                if lista_de_id[15] == 396:
                    lista_de_id[15] = 1000042
                if lista_de_id[15] == 69:
                    lista_de_id[15] = 1000041
                if lista_de_id[15] == 5216:
                    lista_de_id[15] = 1000040

        colunas = [desc[0] for desc in cursor.description]
        dadosTodasEmpresas=[]
        for lista in resultados:
            for item in lista:
                dadosTodasEmpresas.append(item)

        df_badnetPersas = pd.DataFrame.from_records(dadosTodasEmpresas, columns=colunas)

        #Busca os as versões de upload da empresa no servidor
        sSQLVersao = """ 
            Select ver.idagente,ver.versao, ver.DatadeRegistro 
		        FROM [SGT_DEV].[BADNET].[tabelaComentarios] ver
			    inner join (select idagente, datarevrea,max([DatadeRegistro]) as DatadeRegistro
			    FROM [SGT_DEV].[BADNET].[tabelaComentarios]        
			    WHERE versao < 100 and 
                (idagente = ? and year (dataRevrea) = ?) 
                group by idagente,datarevrea) tab on tab.idagente=ver.idagente and tab.datarevrea=ver.datarevrea and tab.dataderegistro=ver.dataderegistro           
                ORDER BY 1  
            """
        #Construção do dataframe das verões de upload das empresas
        resultadosV=[]
        for empresa in empresaPD:   
            cursor.execute(sSQLVersao, (int(empresa), int(ano)))
            resultados_empresa = cursor.fetchall()
            resultadosV.append(resultados_empresa)

        colunasV = [desc[0] for desc in cursor.description]
        dadosTodasVersoes=[]
        for listaV in resultadosV:
            for itemV in listaV:
                dadosTodasVersoes.append(itemV)

        df_badnetPersasV = pd.DataFrame.from_records(dadosTodasVersoes, columns=colunasV)

        #Tabela de Financeiros
        SQL_Financeiros = """SELECT 
          [sigla]
          ,year ([dataRevrea]) as Ano
          ,[IdAgente]
          ,[idSREAg]
          ,[id_natureza_dado]
          ,[id_detalhe_natureza]               
          ,[idDet1]
          ,[idDet2]
          ,[idDet3]
          ,[id_grupo_tarifa]
          ,[id_detalhe_grupo_tarifa]
          ,[id_tipo_tarifa]
          ,[id_subgrupo] 
          ,[id_posto]
          ,[id_unidade]
          ,[id_UC]
          ,[valor]                             
          ,[DatadeRegistro]
        FROM [SGT_DEV].[BADNET].[BADNETFilt]
        WHERE (year(dataRevrea)= ? and month(dataRevrea)= ? and idAgente = ? )
        and((id_Natureza_dado in (2) and id_Detalhe_Natureza in (2) and idDet1 in (22,20,999) and idDet2 in (999999,60,62) and 
        idDet3 in (999,179,178,242,235,241,240,106,220,523,524,525,526,82,81,503,512,73,102,120,123,167,510,518,529,512,234,74,502,515,516,514,180,519,522,527,528,529,530,531,532,533) 
        and id_Grupo_TARIFA IN (99,1,2,3,4) and id_Detalhe_Grupo_TARIFA IN (999,7,11,9,8,10,24,14,15,16,17,18,19,12,1,2,30,28,29,25,3) 
        and id_Tipo_Tarifa IN (9,3) 
        and id_SUBGRUPO IN (99,3) and id_POSTO IN (9) and id_Unidade IN (3) and id_UC IN  (99999) )
            )"""

        Res_Financeiros=[]                                                 #Fazer uma tabela com as tarifas B1 antigas
        for empresa in empresaPD:   
            cursor.execute(SQL_Financeiros, (int(ano), int(mes), int(empresa)))
            Financeiros_empresas = cursor.fetchall()
            Res_Financeiros.append(Financeiros_empresas)

        Financeiros_colunas = [desc[0] for desc in cursor.description]
        dadosRes_Financeiros=[]
        for listaF in Res_Financeiros:
            for itemF in listaF:
                dadosRes_Financeiros.append(itemF)

        df_badnet_Financeiros = pd.DataFrame.from_records(dadosRes_Financeiros, columns=Financeiros_colunas)

        #Tabela de tafifas B1
        SQL_TE_B1 = """SELECT distinct
         Badnet.[sigla]
		,year (badnet.[dataRevrea]) as Ano
        ,badnet.[IdAgente]
        ,badnet.[idSREAg]
        ,badnet.[valor]                             
        ,concat('REH ',NumAto,'/', right(reh1.reh,4)) as Resolucao
        FROM [SGT_DEV].[BADNET].[BADNETFilt] Badnet/*[TabelaBADNET]*/
		inner join  
		(
		select REH_ini.* from [SGT_DEV].[BanTAR].[tblEventos] REH_ini 
		inner join (
		select distinct tab.empresa, max(tab.DataVigencia) as datarevreaMax from  [SGT_DEV].[BanTAR].[tblEventos] tab where yEAR(tab.DataVigencia) in (?,?) 
		group by tab.empresa 
		) maiorData on maiorData.datarevreaMax=reh_ini.DataVigencia and maiordata.empresa=reh_ini.empresa
		) 
		REH1 on REH1.Empresa=Badnet.idagente and REH1.DataVigencia=BADNET.dataRevrea
        WHERE ( idsreag=220 or idsreag<200) and
		(reh1.TipoEvento ='Ordinário' or reh1.TipoEvento is null) and
		
		(id_Natureza_dado in (7) and id_Detalhe_Natureza in (99) 
		and idDet1 in (130) and idDet2 in (999999) and idDet3 in (999) and id_Grupo_TARIFA IN (99)     
		and id_Detalhe_Grupo_TARIFA IN (999) and id_Tipo_Tarifa IN (9) and id_SUBGRUPO IN (99) and id_POSTO IN (9) 
		and id_Unidade IN (5) and id_UC IN  (99999) )
		order by 1"""
        
        Res_TE_B1=[]
        cursor.execute(SQL_TE_B1, (int(ano),int(ano-1),))
        Res_TE_B1=cursor.fetchall()
          
        TE_B1_colunas = [desc[0] for desc in cursor.description]
        df_badnet_TE_B1 = pd.DataFrame.from_records(Res_TE_B1, columns=TE_B1_colunas)


        SQL_TE_B1_Perm = """
        SELECT [IdAgente]
        ,[AnoRef]
        ,[BaseTarifária]
        ,[SUBGRUPO]
        ,[TUSD]
        ,[TE]
        FROM [SGT_DEV].[BanTAR].[TAv]
            where AnoRef = ? and evento = 0 and idagente = ? 
            and BaseTarifária = 'Tarifa de Aplicação' 
            and subgrupo = 'b1' 
            and modalidade = 'convencional' 
            and classe = 'residencial'
            and SUBCLASSE = 'residencial'
            and detalhe = 'não se aplica'"""

        Res_TE_B1_Perm=[]                                                #Fazer uma tabela com as tarifas B1 antigas
        for empresa in empresaPD:   
            cursor.execute(SQL_TE_B1_Perm, (int(ano-1), int(empresa)))
            TE_B1_Perm = cursor.fetchall()
            Res_TE_B1_Perm.append(TE_B1_Perm)

        TE_B1_Perm_colunas = [desc[0] for desc in cursor.description]
        dadosRes_TE_B1_Perm=[]
        for listaTE_Perm in Res_TE_B1_Perm:
            for itemTE_Perm in listaTE_Perm:
                dadosRes_TE_B1_Perm.append(itemTE_Perm)

        df_badnet_TE_B1_Perm = pd.DataFrame.from_records(dadosRes_TE_B1_Perm, columns=TE_B1_Perm_colunas)


        SQL_Pleito = """SELECT 
        year([DataRevRea]) as Ano
        ,[idAgente]
        ,[carta]
        ,[TipoProcesso]
            FROM [SGT_DEV].[BADNET].[InfoPleitoPersas]
            where  year([DataRevRea]) = ? 
            and idagente= ?"""

        Res_Pleito=[]                                                
        for empresa in empresaPD:   
            cursor.execute(SQL_Pleito, (int(ano), int(empresa)))
            Pleito_empresa = cursor.fetchall()
            Res_Pleito.append(Pleito_empresa)

        Res_Pleito_colunas = [desc[0] for desc in cursor.description]
        dadosRes_Pleito=[]
        for listaPleito in Res_Pleito:
            for itemPleito in listaPleito:
                    dadosRes_Pleito.append(itemPleito)


        df_badnet_Pleito = pd.DataFrame.from_records(dadosRes_Pleito, columns=Res_Pleito_colunas)
        #Busca os as versões de upload da empresa no servidor
        
        sSQL_TUSD = """ SELECT bdtar1.[AGENTE_TAR]
	    ,[SAMP_POSTO_TARIFARIO]
	    ,[SAMP_SUBGRUPO]
            ,[AnoRef]
            ,[IdAgente]
            ,[Nao se aplica] as Tarifa
            FROM bantar.bdtar bdtar1 inner join 
			(
			select agente_tar, max(anoref) as maiorAno FROM bantar.bdtar where anoref in ( ?, ?) group by AGENTE_TAR
			) anotab on  Anotab.AGENTE_TAR=bdtar1.AGENTE_TAR and Anotab.maiorAno=bdtar1.AnoRef
            WHERE 
           CodTar LIKE '%050100101%'            
            AND bdtar1.AGENTE_TAR = ?
            AND SAMP_POSTO_TARIFARIO IN (1,2,4)
            AND SAMP_UNIDADE_PRIMARIA IN (2,3)
            AND SAMP_UNIDADE_SECUNDARIA = 2
            AND SAMP_UNIDADE_TERCIARIA IN (2,3)
            AND SAMP_BASE_TARIFA = 5"""
        

        
        
        resultados_empresa_ano_atual = []
        for empresa in empresaPD:  
            cursor.execute(sSQL_TUSD, (int(ano), int(ano-1), int(empresa)))
            resultados_empresa = cursor.fetchall()
            resultados_empresa_ano_atual.append(resultados_empresa)
        
        colunas_TUSD = [desc[0] for desc in cursor.description]

        resultados_TUSD=[]
        for lista in resultados_empresa_ano_atual:   
            for item in lista:
                 resultados_TUSD.append(item)  
                 
        df_badnetPersas_TUSD = pd.DataFrame.from_records(resultados_TUSD, columns=colunas_TUSD)

        #Desconecta do servidor
        conn.close()

        #Cola os dataframes na planilha resumo.
        wk=xw.books.open(caminho_arquivo)

        #Iniciar o aplicativo Excel	
        app = xw.App()

        #Cola os dataframe na planilha
        planilhaBD2 = wk.sheets("BD2")
        planilhaBD2.activate()
        planilhaBD2.range('A:T').options(index=True).value = df_badnetPersas
        planilhaBD2.range('AA:AH').value = df_badnetPersasV
        planilhaBD2.range('AG:AY').options(index=False).value = df_badnet_Financeiros
        planilhaBD2.range('BB1:BH130').options(index=False).value = df_badnet_TE_B1
        planilhaBD2.range('BI1:BN130').options(index=False).value = df_badnet_TE_B1_Perm
        planilhaBD2.range('BQ1:BT130').options(index=False).value = df_badnet_Pleito	
        planilhaBD2.range('BW1:BZ150').options(index=False).value = df_badnetPersas_TUSD

        # Fechar o aplicativo Excel
        app.quit()

