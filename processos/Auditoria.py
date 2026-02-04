import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import json
import sys

def executar_auditoria():

    # Detecta se o script está empacotado (PyInstaller)
    if getattr(sys, 'frozen', False):
        caminho = Path(sys.executable)
    else:
        caminho = Path(__file__).resolve()

    diretorio = caminho.parent

    # Recebe datas
    data_inicial_str = input("Digite a data inicial (dd/mm/yyyy): ")
    data_final_str = input("Digite a data final (dd/mm/yyyy): ")

    try:
        data_inicial = datetime.strptime(data_inicial_str, "%d/%m/%Y")
        data_final = datetime.strptime(data_final_str, "%d/%m/%Y")

        inicio_mes = pd.Timestamp(data_inicial)
        fim_mes = pd.Timestamp(data_final)

    except ValueError:
        print("Formato de data inválido! Use o formato dd/mm/yyyy.")
        exit(1)

    dados = {
        "data_inicial": data_inicial_str,
        "data_final": data_final_str
    }
        
    #cria o data.json com a variável da DATA INICIO e DATA FIM

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False)

    print('Gerando auditoria, aguarde!')


    # Leitura de arquivos usando caminhos dinâmicos
    caminhoBi00 = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\planilhas\Auxiliares\00 - Relação de Lotação.csv"
    caminhoFrequencias_seap = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\planilhas\Auxiliares\Consulta frequências dos funcionários.csv"
    caminhoP1 = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\planilhas\SEDUCAL.xlsx"
    caminhoP2 = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\planilhas\SEDUCMZ.xlsx"


    

    bi00 = pd.read_csv(caminhoBi00, sep=';')
    frequencias_seap = pd.read_csv(caminhoFrequencias_seap, sep=';')
    p1 = pd.read_excel(caminhoP1)
    p2 = pd.read_excel(caminhoP2)

    # <--- CORREÇÃO ADICIONADA --->
    # Padroniza os nomes das colunas para garantir que sejam idênticos antes de concatenar.
    # Remove espaços extras no início/fim e substitui espaços no meio por underscores (_).
    # Isso resolve o problema de colunas duplicadas como 'SALDOFALTANTE'.
    p1.columns = p1.columns.str.strip().str.replace(' ', '_', regex=False)
    p2.columns = p2.columns.str.strip().str.replace(' ', '_', regex=False)
    # <--- FIM DA CORREÇÃO --->

    base_path = diretorio / "Faltas"

    bi00 = bi00[bi00['TpLotacao'] != 'PREFEITURA MUNICIPAL']
    bi00 = bi00.melt(
        id_vars=['DRE', 'Municipio', 'Lotacao'],
        value_vars=['CodLotacao_Seap', 'Cod_Lotacao'],
        value_name='COD SETOR'
    )[['COD SETOR', 'DRE', 'Municipio', 'Lotacao']]
    bi00['COD SETOR'] = bi00['COD SETOR'].astype(int)
    bi00.drop_duplicates(subset=['COD SETOR'], keep='first', inplace=True)


    frequencias_seap['dtini'] = pd.to_datetime(frequencias_seap['dtini'], dayfirst=True, errors='coerce')
    frequencias_seap = frequencias_seap[
        frequencias_seap['frequencia'].isin([
            '8 - FALTA INJUSTIFICADA',
            '34 - FALTA INJUSTIFICADA PROPORCIONAL'
        ])
    ].copy()
    frequencias_seap['qtd'] = np.where(
        frequencias_seap['frequencia'] == '34 - FALTA INJUSTIFICADA PROPORCIONAL',
        frequencias_seap['quantidade'] / 60,
        frequencias_seap['quantidade']
    )
    frequencias_seap = frequencias_seap[['frequencia', 'numfunc', 'numvinc', 'dtini', 'qtd']]
    frequencias_seap.drop_duplicates(subset=['numfunc', 'numvinc', 'dtini'], keep='first', inplace=True)
    frequencias_seap['consta seap'] = 'SIM'


    base = pd.concat([p1, p2])
    base['DataFrequencia'] = pd.to_datetime(base['DataFrequencia'], dayfirst=True, errors='coerce')
    base = base[base['MF'].str.contains('INJUSTIFICADA', na=False)].copy()

    base = base [
        (base['DataFrequencia'] >= inicio_mes) &
        (base['DataFrequencia'] <= fim_mes)
    ]

    # <--- ALTERAÇÃO SUTIL AQUI --->
    # Garante que estamos selecionando a coluna com o nome padronizado (caso ela tivesse espaços)
    # Se o nome original for 'SALDO FALTANTE', aqui ficaria 'SALDOFALTANTE'.
    # No seu caso, o nome parece não ter espaço, então 'SALDOFALTANTE' continua igual.
# Garante que estamos selecionando a coluna com o nome padronizado...
    base = base[['SETOR', 'NOME', 'MATRICULA', 'VINCULO', 'MF', 'DataFrequencia', 'HORAEXECUTADA','SALDOFALTANTE']]

    # --- INSERÇÃO DA LIMPEZA DE MATRÍCULA E VÍNCULO (ESTRATÉGIA INT) ---
    
    # Objetivo: Garante que todos os IDs (ex: 12345.0 ou '00123') virem INT limpos (12345)
    
    # 1. Limpa Matrícula e Vínculo (base)
    base['MATRICULA'] = pd.to_numeric(base['MATRICULA'], errors='coerce').fillna(0).astype(int)
    base['VINCULO'] = pd.to_numeric(base['VINCULO'], errors='coerce').fillna(0).astype(int)
    
    # 2. Limpa Matrícula e Vínculo (frequencias_seap)
    frequencias_seap['numfunc'] = pd.to_numeric(frequencias_seap['numfunc'], errors='coerce').fillna(0).astype(int)
    frequencias_seap['numvinc'] = pd.to_numeric(frequencias_seap['numvinc'], errors='coerce').fillna(0).astype(int)
    
    # --- FIM DA LIMPEZA CONSISTENTE ---
    escolas = base[base['SETOR'].str.match(r'^\d')].copy()
    escolas[['COD SETOR', 'DESC SETOR']] = escolas['SETOR'].str.split('-', n=1, expand=True)
    escolas['COD SETOR'] = escolas['COD SETOR'].astype(int)

    escolas = pd.merge(escolas, bi00, left_on='COD SETOR', right_on='COD SETOR', how='left')
    escolas = pd.merge(escolas, frequencias_seap, left_on=['MATRICULA', 'VINCULO', 'DataFrequencia'], right_on=['numfunc', 'numvinc', 'dtini'], how='left')
    escolas = escolas[escolas['consta seap'].isna()]
    escolas['UNIDADE ESCOLAR'] = escolas['COD SETOR'].astype(str) + ' - ' + escolas['DESC SETOR']
    escolas = escolas[['DRE', 'Municipio', 'UNIDADE ESCOLAR', 'NOME', 'MATRICULA', 'VINCULO', 'MF', 'DataFrequencia', 'SALDOFALTANTE']]
    escolas.drop_duplicates(inplace=True)

    outros = base[~base['SETOR'].str.match(r'^\d')].copy()
    outros = pd.merge(outros, frequencias_seap, left_on=['MATRICULA', 'VINCULO', 'DataFrequencia'], right_on=['numfunc', 'numvinc', 'dtini'], how='left')
    outros = outros[outros['consta seap'].isna()]

    def atribuir_dre(setor):
        setor = setor.upper()
        if "PRIMAVERA DO LESTE" in setor:
            return "DRE PRIMAVERA DO LESTE"
        elif "CUIABA" in setor or "CUIABÁ" in setor or "VARZEA GRANDE" in setor:
            return "DRE METROPOLITANA"
        elif "ALTA FLORESTA" in setor:
            return "DRE ALTA FLORESTA"
        elif "CONFRESA" in setor:
            return "DRE CONFRESA"
        elif "CACERES" in setor or "CÁCERES" in setor:
            return "DRE CACERES"
        elif "PONTES E LACERDA" in setor or "DREPLC" in setor:
            return "DRE PONTES E LACERDA"
        elif "TANGARÁ DA SERRA" in setor or "TANGARA DA SERRA" in setor:
            return "DRE TANGARA DA SERRA"
        elif "RONDONOPOLIS" in setor or "RONDONÓPOLIS" in setor:
            return "DRE RONDONOPOLIS"
        elif "QUERENCIA" in setor or "BARRA DO GARÇAS" in setor or "DREBG" in setor or "BARRA DO GARCAS" in setor:
            return "DRE BARRA DO GARCAS"
        elif "DREDIAM" in setor or "DIAMANTINO" in setor:
            return "DRE DIAMANTINO"
        elif "JUINA" in setor or "DREJUI" in setor:
            return "DRE JUINA"
        elif "MATUPÁ" in setor or "MATUPA" in setor:
            return "DRE MATUPA"
        elif "SINOP" in setor or "DRESNP" in setor:
            return "DRE SINOP"
        else:
            return "ORGAO CENTRAL"
        



    outros["DRE"] = outros["SETOR"].apply(atribuir_dre)

    outros['SETOR'] = outros['SETOR'].str.strip()
    outros = outros[['DRE','Municipio', 'SETOR', 'NOME', 'MATRICULA', 'VINCULO', 'MF', 'DataFrequencia', 'SALDOFALTANTE']]
    outros.drop_duplicates(inplace=True)


    


    # --- INÍCIO DO NOVO BLOCO DE SALVAMENTO CONSOLIDADO ---

    # PASSO 2: Criar a Lista Mestra de DREs
    lista_dres_escolas = escolas['DRE'].dropna().unique()
    lista_dres_outros = outros['DRE'].dropna().unique()
    lista_mestra_dres = np.union1d(lista_dres_escolas, lista_dres_outros)

    print("Iniciando a criação do arquivo Excel consolidado...")
    print(lista_mestra_dres)
    print(f'Total: {len(lista_mestra_dres)}')

    # PASSO 3: Abrir o "Arquivo Excel Mestre"
    nome_arquivo_final = "Auditoria_Completa_por_DRE.xlsx"
    caminho_arquivo_final = base_path / nome_arquivo_final
    base_path.mkdir(parents=True, exist_ok=True) # Garante que a pasta FALTAS exista

    print(f"Salvando relatório em: {caminho_arquivo_final}")

    total_abas_criadas=0

    # Abre o "escritor" de Excel. O bloco 'with' garante que ele será salvo e fechado no final.
    with pd.ExcelWriter(caminho_arquivo_final, engine='openpyxl') as writer:
        
        # PASSO 4: Loop Principal por cada DRE
        for dre_nome in lista_mestra_dres:
            
            # PASSO 5.1: Filtrar e salvar a aba de ESCOLAS para a DRE atual
            dados_escola_dre = escolas[escolas['DRE'] == dre_nome]
            
            # Verifica se existem dados de escolas para esta DRE antes de salvar
            if not dados_escola_dre.empty:
                # O Excel tem um limite de 31 caracteres para nomes de abas
                nome_aba_escola = dre_nome.replace("DRE", "ESCOLA")[:31] 
                dados_escola_dre.to_excel(writer, sheet_name=nome_aba_escola, index=False)


                print(f"  > [ESCOLA] Aba criada: {nome_aba_escola}")  # <--- COLE AQUI
                total_abas_criadas += 1
            
            # PASSO 5.2: Filtrar e salvar a aba de DRE (Administrativo) para a DRE atual
            dados_outros_dre = outros[outros['DRE'] == dre_nome]
            
            # Verifica se existem dados administrativos para esta DRE antes de salvar
            if not dados_outros_dre.empty:
                nome_aba_adm = dre_nome[:31]
                dados_outros_dre.to_excel(writer, sheet_name=nome_aba_adm, index=False)


                print(f"  > [ADM]    Aba criada: {nome_aba_adm}")     # <--- COLE AQUI
                total_abas_criadas += 1

    # --- FIM DO NOVO BLOCO DE SALVAMENTO ---




    print(f'Relatório finalizado. Total de {total_abas_criadas} foram criadas')
    print('Auditoria concluída com sucesso!')