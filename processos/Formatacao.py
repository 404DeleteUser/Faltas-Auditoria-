import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import os

def executar_formatacao():
    # ... resto do código

    # --- CONFIGURAÇÃO OBRIGATÓRIA ---

    # 1. Coloque aqui o caminho para o seu arquivo Excel
    #    Exemplo Windows: r"C:\usuarios\seu_nome\documentos\minha_planilha.xlsx"
    #    Exemplo macOS/Linux: "/home/seu_nome/documentos/minha_planilha.xlsx"
    CAMINHO_DO_ARQUIVO_EXCEL = r"C:\Users\08477936137\Downloads\AuditoriaGit\processos\FALTAS\Auditoria_Completa_por_DRE.xlsx"

    # 2. Definição das cores (sem o '#')
    COR_CABECALHO = '356854'       # Verde escuro para o fundo do cabeçalho
    COR_TEXTO_CABECALHO = 'FFFFFF' # Branco para o texto do cabeçalho
    COR_LINHA_BRANCA = 'FFFFFF'    # Branco explícito para a primeira linha de dados
    COR_LINHA_ZEBRADA = 'beebda'   # Verde claro para as linhas alternadas

    # --- FIM DA CONFIGURAÇÃO ---


    def auto_ajustar_colunas(ws):
        """
        Função auxiliar para ajustar a largura das colunas com base no conteúdo.
        NOTA: Esta é uma aproximação, pois o openpyxl não calcula a largura 
        exata de renderização da fonte como o Excel faz.
        """
        for col_idx in range(1, ws.max_column + 1):
            column_letter = get_column_letter(col_idx)
            max_length = 0
            
            try:
                for row_idx in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value:
                        # Ajusta o tamanho máximo baseado no comprimento do texto
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                
                # Adiciona um pequeno preenchimento (padding)
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column_letter].width = adjusted_width
                
            except Exception as e:
                print(f"  [Aviso] Não foi possível ajustar coluna {column_letter}: {e}")


    def formatar_planilha_excel(caminho_arquivo):
        """
        Função principal que abre um arquivo Excel e formata todas as suas abas.
        """
        if not os.path.exists(caminho_arquivo):
            print(f"ERRO: O arquivo não foi encontrado em: {caminho_arquivo}")
            print("Por favor, preencha a variável CAMINHO_DO_ARQUIVO_EXCEL corretamente.")
            return

        print(f"Abrindo o arquivo: {caminho_arquivo}...")
        
        try:
            # Carrega o arquivo Excel
            wb = openpyxl.load_workbook(caminho_arquivo)
            
            # Define os estilos de preenchimento e fonte (fazemos isso fora do loop por eficiência)
            fill_cabecalho = PatternFill(start_color=COR_CABECALHO, end_color=COR_CABECALHO, fill_type="solid")
            font_cabecalho = Font(color=COR_TEXTO_CABECALHO, bold=True)
            
            fill_branca = PatternFill(start_color=COR_LINHA_BRANCA, end_color=COR_LINHA_BRANCA, fill_type="solid")
            fill_zebrada = PatternFill(start_color=COR_LINHA_ZEBRADA, end_color=COR_LINHA_ZEBRADA, fill_type="solid")

            print(f"Iniciando a formatação para {len(wb.worksheets)} aba(s)...")

            # Itera por todas as abas (worksheets) do arquivo
            for ws in wb.worksheets:
                print(f"Formatando a aba: \"{ws.title}\"...")
                
                max_linha = ws.max_row
                max_coluna = ws.max_column

                if max_linha == 0 or max_coluna == 0:
                    print(f"  A aba \"{ws.title}\" está vazia. Nenhuma formatação aplicada.")
                    continue

                # 1. Formata o Cabeçalho (Linha 1)
                for col_idx in range(1, max_coluna + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.fill = fill_cabecalho
                    cell.font = font_cabecalho

                # 2. Aplica as cores zebradas (a partir da Linha 2)
                for row_idx in range(2, max_linha + 1):
                    # Decide qual cor usar
                    # Linha 2 (par) = branca, Linha 3 (ímpar) = zebrada, Linha 4 (par) = branca...
                    current_fill = fill_branca if row_idx % 2 == 0 else fill_zebrada
                    
                    for col_idx in range(1, max_coluna + 1):
                        ws.cell(row=row_idx, column=col_idx).fill = current_fill

                # 3. Congela a primeira linha
                # 'A2' congela a linha 1
                ws.freeze_panes = 'A2'
                
                # 4. Auto-ajusta as colunas (aproximação)
                auto_ajustar_colunas(ws)

            wb.save(caminho_arquivo)
            print("\nFormatação concluída com sucesso!")

        except Exception as e:
            print(f"Ocorreu um erro durante a execução: {e}")
            if isinstance(e, PermissionError):
                print("ERRO: Você precisa fechar o arquivo Excel antes de executar o script.")

