import pandas as pd
import numpy as np
import json
from datetime import datetime
from pathlib import Path
import sys

# --- 1. CONFIGURAÇÃO DOS CAMINHOS ---
licencas = {
    "LicencaFerias": r"C:\Users\08477936137\Downloads\AuditoriaGit\planilhas\BD_Licencas\Ferias.csv",
    "LicencaLIP": r"C:\Users\08477936137\Downloads\AuditoriaGit\planilhas\BD_Licencas\LIP.xlsx",
    # Adicione outras aqui se precisar
}

# --- 2. REGRAS DE TRADUÇÃO (DE-PARA) ---
colunas_licencas = {
    "LicencaFerias": {
        "sep": ";", 
        "renomear": {
            'numfunc': 'MATRICULA',
            'numvinc': 'VINCULO',
            'dtini': 'DATA_INICIO',
            'dtfim': 'DATA_FIM'
        },
        "nome_na_planilha": "FÉRIAS"
    },
    "LicencaLIP": {
        "sep": ";", 
        "renomear": {
            'numfunc': 'MATRICULA',
            'numvinc': 'VINCULO',
            'dtini': 'DATA_INICIO',
            'dtfim': 'DATA_FIM'
        },
        "nome_na_planilha": "LICENÇA PRÊMIO"
    }
}

# Caminho do arquivo alvo (gerado pelo Auditoria.py)
Auditoria = r"C:\Users\08477936137\Downloads\AuditoriaGit\processos\Faltas\Auditoria_Completa_por_DRE.xlsx"

def auditoria_sequencial():
    print("\n=== INICIANDO AUDITORIA SEQUENCIAL DE LICENÇAS ===")
    
    caminho_auditoria = Auditoria
    print(f"Carregando arquivo base: {caminho_auditoria}")

    # 1. CARREGAR A AUDITORIA (TODAS AS ABAS)
    try:
        dict_auditoria = pd.read_excel(caminho_auditoria, sheet_name=None)
    except FileNotFoundError:
        print("ERRO: Arquivo de Auditoria não encontrado. Rode o Auditoria.py primeiro.")
        return

    # 2. PREPARAR TODAS AS ABAS DA AUDITORIA
    print(f"Preparando {len(dict_auditoria)} abas para verificação...")
    
    for nome_aba, df_auditoria in dict_auditoria.items():
        # Cria a coluna de status se não existir
        if 'Status_Licenca' not in df_auditoria.columns:
            df_auditoria['Status_Licenca'] = ''

        # --- PADRONIZAÇÃO PARA INTEIRO (AUDITORIA) ---
        # Isso remove .0, remove espaços e garante que é número puro
        df_auditoria['MATRICULA'] = pd.to_numeric(df_auditoria['MATRICULA'], errors='coerce').fillna(0).astype(int)
        df_auditoria['VINCULO'] = pd.to_numeric(df_auditoria['VINCULO'], errors='coerce').fillna(0).astype(int)
        
        # Padroniza a data
        df_auditoria['DataFrequencia'] = pd.to_datetime(df_auditoria['DataFrequencia'], dayfirst=True, errors='coerce')

    # 3. LOOP PRINCIPAL: UMA LICENÇA POR VEZ
    for nome_licenca, regras in colunas_licencas.items():
        
        caminho_arquivo = licencas.get(nome_licenca) # Pega o caminho usando a chave certa

        if not caminho_arquivo:
            print(f"Pulando {nome_licenca} (Caminho não configurado)...")
            continue

        print(f"\n>>> Processando Licença: {nome_licenca}")

        try:
            # A. Carrega o arquivo de licença
            if caminho_arquivo.endswith('.xlsx'):
                df_licenca = pd.read_excel(caminho_arquivo)
            else:
                df_licenca = pd.read_csv(caminho_arquivo, sep=regras['sep'], on_bad_lines='skip')




            # --- INSERIR O DIAGNÓSTICO AQUI ---
            print(f"--- DIAGNÓSTICO: CABEÇALHOS BRUTOS DO {nome_licenca} ---")
            print(df_licenca.columns.tolist())
            print("---------------------------------------------------\n")

            # B. Limpa e Padroniza a Licença
            df_licenca.columns = df_licenca.columns.str.strip().str.replace(' ', '_', regex=False)
            df_licenca.rename(columns=regras['renomear'], inplace=True)







            # B. Limpa e Padroniza a Licença
            df_licenca.columns = df_licenca.columns.str.strip().str.replace(' ', '_', regex=False)
            df_licenca.rename(columns=regras['renomear'], inplace=True)
            

            # DEBUG: Checagem pós-renomeação
            print(f"--- DEBUG: Dados de {nome_licenca} ---")
            print(f"Total de registros: {len(df_licenca)}")
            
            if not df_licenca.empty:
                print(f"Cabeçalhos (após Renomear): {df_licenca.columns.tolist()}")
                
                # Para fins de diagnóstico, vou tentar imprimir o dtype (tipo)
                # O .dtype só funciona se a coluna existir, então vamos testar
                try:
                    print(f"  Tipo MATRICULA (Esperado INT): {df_licenca['MATRICULA'].dtype}")
                except KeyError:
                    # Se der erro aqui, a coluna 'MATRICULA' NÃO EXISTE (o rename falhou)
                    print("  🚨 ERRO CRÍTICO: A coluna 'MATRICULA' NÃO FOI ENCONTRADA.")
                    print(f"  Colunas que existem: {df_licenca.columns.tolist()}")
                    
                # Mostra o conteúdo da primeira linha
                print("\nPrimeira linha (RAW DATA) - Matrícula e Vínculo:")
                print(df_licenca[['MATRICULA', 'VINCULO', 'DATA_INICIO', 'DATA_FIM']].head(1).to_string(index=False))

            print("---------------------------------------------------\n")


            # --- PADRONIZAÇÃO PARA INTEIRO (LICENÇA) ---
            # Aplica a mesma lógica da auditoria para garantir o "match"
            df_licenca['MATRICULA'] = pd.to_numeric(df_licenca['MATRICULA'], errors='coerce').fillna(0).astype(int)
            df_licenca['VINCULO'] = pd.to_numeric(df_licenca['VINCULO'], errors='coerce').fillna(0).astype(int)
            
            # Padroniza as datas
            df_licenca['DATA_INICIO'] = pd.to_datetime(df_licenca['DATA_INICIO'], dayfirst=True, errors='coerce')
            df_licenca['DATA_FIM'] = pd.to_datetime(df_licenca['DATA_FIM'], dayfirst=True, errors='coerce')

            tipo_da_licenca = regras['nome_na_planilha']
            count_encontrados = 0

            # C. LOOP INTERNO: VERIFICA EM CADA ABA DA AUDITORIA
            for nome_aba, df_auditoria in dict_auditoria.items():
                
                # Itera sobre cada falta nesta aba
                for idx, linha_falta in df_auditoria.iterrows():
                    
                    # Se já tem justificativa, pula
                    if pd.notna(linha_falta['Status_Licenca']) and str(linha_falta['Status_Licenca']) != '':
                        continue

                    # Dados da Falta (Já convertidos para INT lá em cima)
                    mat = linha_falta['MATRICULA']
                    vinc = linha_falta['VINCULO']
                    data_falta = linha_falta['DataFrequencia']

                    # Se matrícula for 0 (inválida), pula
                    if mat == 0: continue

                    # Filtra a licença para esse funcionário (INT com INT bate perfeito)
                    # Usamos query para ser rápido
                    licencas_match = df_licenca.query("MATRICULA == @mat and VINCULO == @vinc")

                    # Verifica as datas
                    for _, lic in licencas_match.iterrows():
                        if lic['DATA_INICIO'] <= data_falta <= lic['DATA_FIM']:
                            
                            # ACHOU!
                            df_auditoria.at[idx, 'Status_Licenca'] = tipo_da_licenca
                            count_encontrados += 1
                            
                            # Print de Confirmação
                            print(f"   [ACHOU!] {nome_aba} | Mat:{mat} | {tipo_da_licenca} em {data_falta.strftime('%d/%m/%Y')}")
                            break # Para de procurar nesta licença
            
            print(f"   -> Total encontrados em {nome_licenca}: {count_encontrados}")

        except Exception as e:
            print(f"ERRO CRÍTICO ao processar {nome_licenca}: {e}")

    # 4. SALVAR O RESULTADO FINAL
    print("\nSalvando arquivo verificado...")
    caminho_saida = caminho_auditoria
    
    try:
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            for nome_aba, df_final in dict_auditoria.items():
                # Formata a data para ficar bonita no Excel final
                if 'DataFrequencia' in df_final.columns:
                    df_final['DataFrequencia'] = df_final['DataFrequencia'].dt.strftime('%d/%m/%Y')
                
                df_final.to_excel(writer, sheet_name=nome_aba, index=False)
        
        print(f"SUCESSO! Relatório salvo em: {caminho_saida}")
        
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")
