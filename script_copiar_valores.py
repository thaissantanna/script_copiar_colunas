import pandas as pd
import os


class ProcessadorExcel:
    def __init__(self, arquivo_entrada, arquivo_saida, mapeamento_colunas, filtros, coluna_multiplicar=None, fator_multiplicacao=None, aba_origem="Sheet1", aba_destino="Sheet1"):
        """
        :param arquivo_entrada: Caminho do arquivo Excel de entrada.
        :param arquivo_saida: Caminho do arquivo Excel de saída.
        :param mapeamento_colunas: Dicionário mapeando colunas de origem para destino.
        :param filtros: Dicionário de filtros (ex: {"X": "revenda"}).
        :param coluna_multiplicar: Nome da coluna a ser multiplicada.
        :param fator_multiplicacao: Valor pelo qual a coluna será multiplicada.
        :param aba_origem: Nome da aba de onde os dados serão lidos.
        :param aba_destino: Nome da aba onde os dados serão salvos.
        """
        self.arquivo_entrada = arquivo_entrada
        self.arquivo_saida = arquivo_saida
        # Normaliza nomes de colunas
        self.mapeamento_colunas = {
            k.lower(): v for k, v in mapeamento_colunas.items()}
        self.filtros = {k.lower(): v for k, v in filtros.items()
                        }  # Normaliza filtros
        self.coluna_multiplicar = coluna_multiplicar.lower() if coluna_multiplicar else None
        self.fator_multiplicacao = fator_multiplicacao
        self.aba_origem = aba_origem
        self.aba_destino = aba_destino

    def definir_abas(self, origem, destino):
        """Define as abas de leitura e escrita."""
        self.aba_origem = origem
        self.aba_destino = destino

    def ler_excel(self):
        """Lê o arquivo Excel e aplica os filtros."""
        try:
            df = pd.read_excel(self.arquivo_entrada,
                               sheet_name=self.aba_origem)

            # Normaliza os nomes das colunas
            df.columns = [col.lower().strip() if pd.notna(
                col) else None for col in df.columns]
            # Remove colunas totalmente vazias
            df = df.dropna(axis=1, how='all')

            # Verifica se as colunas existem
            for coluna in self.mapeamento_colunas.keys():
                if coluna not in df.columns:
                    raise ValueError(
                        f"Coluna '{coluna}' não encontrada no arquivo.")

            # Aplica os filtros
            for coluna, valor in self.filtros.items():
                df = df[df[coluna] == valor]

            # Aplica multiplicação, se especificado
            if self.coluna_multiplicar and self.fator_multiplicacao is not None:
                if self.coluna_multiplicar in df.columns:
                    df[self.coluna_multiplicar] *= self.fator_multiplicacao
                else:
                    print(f"Aviso: Coluna '{
                          self.coluna_multiplicar}' para multiplicação não encontrada.")

            # Renomeia as colunas conforme mapeamento
            df = df[list(self.mapeamento_colunas.keys())].rename(
                columns=self.mapeamento_colunas)

            return df
        except Exception as e:
            print(f"Erro ao ler o arquivo Excel: {e}")
            return None

    def salvar_excel(self, df):
        """Salva os dados no arquivo de saída, empilhando se já existir."""
        try:
            with pd.ExcelWriter(self.arquivo_saida, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name=self.aba_destino, index=False)
            print(f"Arquivo salvo na aba '{
                  self.aba_destino}' do arquivo {self.arquivo_saida}")
        except Exception as e:
            print(f"Erro ao salvar o arquivo Excel: {e}")

    def processar(self):
        """Executa o processamento do arquivo."""
        df = self.ler_excel()
        if df is not None:
            self.salvar_excel(df)


# Exemplo de uso
if __name__ == "__main__":
    arquivo_entrada = r"C:\\Users\\SeuUsuario\\Documents\\entrada.xlsx"
    arquivo_saida = r"C:\\Users\\SeuUsuario\\Documents\\saida.xlsx"
    # Exemplo de mapeamento de colunas
    mapeamento_colunas = {"X": "A", "Y": "B", "Z": "C"}
    filtros = {"X": "revenda"}  # Exemplo de filtro
    coluna_multiplicar = "Z"
    fator_multiplicacao = 2

    processador = ProcessadorExcel(arquivo_entrada, arquivo_saida, mapeamento_colunas, filtros,
                                   coluna_multiplicar, fator_multiplicacao, aba_origem="Dados", aba_destino="Resultado")
    processador.processar()
