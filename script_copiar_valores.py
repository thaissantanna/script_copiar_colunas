import pandas as pd
import os

class ProcessadorExcel:
    def __init__(self, arquivo_entrada, arquivo_saida, mapeamento_colunas, filtros, coluna_multiplicar=None, fator_multiplicacao=None):
        """
        :param arquivo_entrada: Caminho do arquivo Excel de entrada.
        :param arquivo_saida: Caminho do arquivo Excel de saída.
        :param mapeamento_colunas: Dicionário mapeando colunas de origem para destino.
        :param filtros: Dicionário de filtros (ex: {"X": "revenda"}).
        :param coluna_multiplicar: Nome da coluna a ser multiplicada.
        :param fator_multiplicacao: Valor pelo qual a coluna será multiplicada.
        """
        self.arquivo_entrada = arquivo_entrada
        self.arquivo_saida = arquivo_saida
        self.mapeamento_colunas = {k.lower(): v for k, v in mapeamento_colunas.items()}  # Normaliza nomes de colunas
        self.filtros = {k.lower(): v for k, v in filtros.items()}  # Normaliza filtros
        self.coluna_multiplicar = coluna_multiplicar.lower() if coluna_multiplicar else None
        self.fator_multiplicacao = fator_multiplicacao

    def ler_excel(self):
        """Lê o arquivo Excel e aplica os filtros."""
        try:
            df = pd.read_excel(self.arquivo_entrada)
            
            # Normaliza os nomes das colunas (removendo "Unnamed" e tratando maiúsculas/minúsculas)
            df.columns = [col.lower().strip() if pd.notna(col) else None for col in df.columns]
            df = df.dropna(axis=1, how='all')  # Remove colunas totalmente vazias
            
            # Verifica se as colunas existem no arquivo
            for coluna in self.mapeamento_colunas.keys():
                if coluna not in df.columns:
                    raise ValueError(f"Coluna '{coluna}' não encontrada no arquivo.")
            
            # Aplica os filtros
            for coluna, valor in self.filtros.items():
                df = df[df[coluna] == valor]
            
            # Aplica multiplicação, se especificado
            if self.coluna_multiplicar and self.fator_multiplicacao is not None:
                if self.coluna_multiplicar in df.columns:
                    df[self.coluna_multiplicar] *= self.fator_multiplicacao
                else:
                    print(f"Aviso: Coluna '{self.coluna_multiplicar}' para multiplicação não encontrada.")
            
            # Renomeia as colunas conforme mapeamento
            df = df[list(self.mapeamento_colunas.keys())].rename(columns=self.mapeamento_colunas)
            
            return df
        except Exception as e:
            print(f"Erro ao ler o arquivo Excel: {e}")
            return None

    def salvar_excel(self, df):
        """Salva os dados no arquivo de saída, empilhando se já existir."""
        try:
            if os.path.exists(self.arquivo_saida):
                df_existente = pd.read_excel(self.arquivo_saida)
                df = pd.concat([df_existente, df], ignore_index=True)
                versao = 1
                novo_arquivo = self.arquivo_saida.replace(".xlsx", f"_v{versao}.xlsx")
                while os.path.exists(novo_arquivo):
                    versao += 1
                    novo_arquivo = self.arquivo_saida.replace(".xlsx", f"_v{versao}.xlsx")
                df.to_excel(novo_arquivo, index=False)
                print(f"Arquivo atualizado e salvo como {novo_arquivo}")
            else:
                df.to_excel(self.arquivo_saida, index=False)
                print(f"Arquivo salvo como {self.arquivo_saida}")
        except Exception as e:
            print(f"Erro ao salvar o arquivo Excel: {e}")

    def processar(self):
        """Executa o processamento do arquivo."""
        df = self.ler_excel()
        if df is not None:
            self.salvar_excel(df)

# Exemplo de uso
if __name__ == "__main__":
    arquivo_entrada = r"C:\Users\SeuUsuario\Documents\entrada.xlsx"
    arquivo_saida = r"C:\Users\SeuUsuario\Documents\saida.xlsx"
    mapeamento_colunas = {"NOME DAS COLUNAS AQUI"} # Nome das colunas "X": "A", "Y": "B"...
    filtros = {}  # Filtra apenas valores "revenda" na coluna X
    coluna_multiplicar = "SE FOR PRA MULTIPLICAR ALGUM VALOR, SE NÃO DEIXAR VAZIO"  # Coluna a ser multiplicada
    fator_multiplicacao = 0  # Fator de multiplicação
    
    processador = ProcessadorExcel(arquivo_entrada, arquivo_saida, mapeamento_colunas, filtros, coluna_multiplicar, fator_multiplicacao)
    processador.processar()