import pandas as pd
import os

# ================= CONFIGURAÇÕES =================
CAMINHO_FALTAS = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\processos\Faltas\Auditoria_Completa_por_DRE.xlsx"
CAMINHO_LICENCAS = r"C:\Users\08477936137\Downloads\AuditoriaDeFaltas\planilhas\Auxiliares\licencas.csv"


## 1. Carregamento de Dados (Não Alterada)
def carregar_dados():
    """Carrega os DataFrames de Faltas (em dict) e Licenças."""
    print("--- [PASSO 1] Carregando bases de dados... ---")
    
    if not os.path.exists(CAMINHO_FALTAS):
        print(f"ERRO: Arquivo de Faltas não encontrado: {CAMINHO_FALTAS}")
        print("Execute o Passo 1 (Auditoria) primeiro.")
        return None, None
        
    if not os.path.exists(CAMINHO_LICENCAS):
        print(f"ERRO: Arquivo de Licenças não encontrado: {CAMINHO_LICENCAS}")
        return None, None

    try:
        # Carrega LICENÇAS
        try:
            df_licencas = pd.read_csv(CAMINHO_LICENCAS, sep=None, engine='python', encoding='latin1')
        except:
            df_licencas = pd.read_csv(CAMINHO_LICENCAS, sep=None, engine='python', encoding='utf-8')
            
        # Carrega FALTAS (Todas as abas)
        dict_faltas = pd.read_excel(CAMINHO_FALTAS, sheet_name=None)
        return dict_faltas, df_licencas
    except Exception as e:
        print(f"Erro crítico ao abrir arquivos: {e}")
        return None, None


## 2. Preparação de Licenças (FUNÇÃO ÚNICA PARA O DF LICENÇAS)
def preparar_licencas(df_licencas):
    """Limpa, padroniza e prepara o DataFrame de Licenças para o cruzamento."""
    print("--- [PASSO 2] Preparando base de licenças... ---")
    
    df_licencas.columns = df_licencas.columns.str.strip()

    # Padroniza 'Func' e 'Vinc' para inteiro
    for dados in ['Func', 'Vinc']:
        df_licencas[dados] = pd.to_numeric(df_licencas[dados], errors='coerce').fillna(0).astype(int)

    # Padroniza DataInicial e DataFinal para tipo data
    for datas in ['DataInicial','DataFinal']:
        df_licencas[datas] = pd.to_datetime(df_licencas[datas], dayfirst=True, errors='coerce') 

    # Renomeia
    df_licencas = df_licencas.rename(columns={'Func': 'MATRICULA', 'Vinc': 'VINCULO'})
    
    # Retorna o DF de licenças pronto
    return df_licencas[['MATRICULA', 'VINCULO', 'Tipo', 'DataInicial', 'DataFinal']]


## 3. Preparação das Faltas (ITERA E PREPARA O DICT DE FALTAS)
def preparar_faltas(dict_faltas):
    """Percorre e padroniza o dicionário de DataFrames de Faltas."""
    print("--- [PASSO 3] Preparando DataFrames de faltas... ---")
    
    dict_faltas_prep = {}
    
    for nome_aba, df in dict_faltas.items():
        # Padroniza nomes de colunas
        df.columns = df.columns.str.strip().str.upper() 
        
        # Padroniza tipos de Matricula/Vinculo
        for servidor_faltas in ['MATRICULA','VINCULO']:
            if servidor_faltas in df.columns:
                df[servidor_faltas] = pd.to_numeric(df[servidor_faltas], errors='coerce').fillna(0).astype(int)
        
        # Padroniza Data de Frequência (nome e tipo)
        cols_data = [c for c in df.columns if 'DATA' in c and 'FREQ' in c]
        nome_col_data = cols_data[0] if cols_data else None

        if nome_col_data:
             df[nome_col_data] = pd.to_datetime(df[nome_col_data], dayfirst=True, errors='coerce')
             # Adiciona a coluna 'Licença'
             df['Licença'] = None
        
        dict_faltas_prep[nome_aba] = df # Adiciona o DataFrame limpo ao novo dicionário
        
    return dict_faltas_prep


## 4. Processar Aba (Não precisa de alteração de lógica, apenas renomeada a variável de licenças)
def processar_aba(df_falta, df_licencas_prep):
    """Cruza uma aba específica de faltas com as licenças."""

    # 1. Encontra o nome da coluna de data de frequência (necessário para o merge)
    cols_data = [c for c in df_falta.columns if 'DATA' in c and 'FREQ' in c]
    nome_col_data = cols_data[0] if cols_data else None

    if not nome_col_data:
        return df_falta 

    if df_falta.empty:
        return df_falta

    # Cria índice temporário (para mapear de volta)
    df_falta = df_falta.reset_index(drop=True)
    df_falta['id_temp'] = df_falta.index

    # 2. Filtro Otimizado
    matriculas_na_aba = df_falta['MATRICULA'].unique()
    licencas_relevantes = df_licencas_prep[df_licencas_prep['MATRICULA'].isin(matriculas_na_aba)]

    # Se ninguém tem licença nessa aba, retorna logo
    if licencas_relevantes.empty:
        return df_falta.drop(columns=['id_temp'])

    # 3. Cruzamento (Merge)
    df_falta_temp = df_falta.rename(columns={nome_col_data: 'DataFreq_Padrao'})
    
    merged = df_falta_temp.merge(
        licencas_relevantes, 
        on=['MATRICULA', 'VINCULO'], 
        how='left'
    )

    # 4. Verifica intervalo de datas
    mask_valida = (
        (merged['DataFreq_Padrao'] >= merged['DataInicial']) & 
        (merged['DataFreq_Padrao'] <= merged['DataFinal'])
    )
    
    licencas_validas = merged[mask_valida]

    # 5. Preenchimento
    if not licencas_validas.empty:
        licencas_unicas = licencas_validas.drop_duplicates(subset=['id_temp']) 
        mapa_licencas = dict(zip(licencas_unicas['id_temp'], licencas_unicas['Tipo']))
        
        df_falta['Licença'] = df_falta['id_temp'].map(mapa_licencas)

    # 6. Limpeza final
    return df_falta.drop(columns=['id_temp'])


## 5. Executar Verificação (FLUXO PRINCIPAL)
def executar_verificacao():
    print("\n=== INICIANDO VERIFICAÇÃO DE LICENÇAS ===")
    
    # 1. Carregamento
    dict_faltas, df_licencas = carregar_dados()
    if dict_faltas is None or df_licencas is None: return

    # 2. Preparação das Licenças (Modular)
    df_licencas_prep = preparar_licencas(df_licencas)

    # 3. Preparação das Faltas (Modular)
    dict_faltas_prep = preparar_faltas(dict_faltas)


    dict_resultado = {}
    total_licencas = 0
    total_abas = len(dict_faltas_prep)

    # 4. Processamento (Iteração)
    for i, (nome_aba, df_aba) in enumerate(dict_faltas_prep.items(), 1):
        print(f" -> [{i}/{total_abas}] Processando aba '{nome_aba}'...", end='\r')
        
        # Verifica se o DF possui as colunas necessárias para o processamento
        colunas = [c for c in df_aba.columns] 
        if not ('MATRICULA' in colunas and 'VINCULO' in colunas and 'Licença' in colunas):
            dict_resultado[nome_aba] = df_aba
            continue
            
        try:
            # Cruza a aba atual com as Licenças preparadas
            df_processado = processar_aba(df_aba, df_licencas_prep) 
            
            # Conta licenças achadas
            total_licencas += df_processado['Licença'].notna().sum()
            
            dict_resultado[nome_aba] = df_processado
            
        except Exception as e:
            print(f"\n[ERRO] Falha na aba '{nome_aba}': {e}")
            dict_resultado[nome_aba] = df_aba

    print(f"\nProcessamento concluído. {total_licencas} licenças identificadas.")

    # 5. Salvamento
    print("Salvando arquivo...")
    try:      
        with pd.ExcelWriter(CAMINHO_FALTAS, engine='openpyxl', date_format='DD/MM/YYYY') as writer:
            for nome_aba, df in dict_resultado.items():
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                
        print(f"Sucesso! Arquivo salvo em: {CAMINHO_FALTAS}")
        
    except PermissionError:
        print("\n[ERRO] Feche o arquivo Excel e tente novamente!")
    except Exception as e:
        print(f"\n[ERRO] Erro ao salvar o arquivo: {e}")

# Chame a função principal no seu main.py
