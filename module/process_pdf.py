import pandas as pd
import tabula


def process_pdf(path: str, pages: str):
    try:
        # Ler as tabelas de todas as páginas indicadas
        tables = tabula.read_pdf(path, pages=pages, lattice=True, multiple_tables=True)

        # Certificar-se de que tabelas foram extraídas
        if not tables or len(tables) == 0:
            return "Nenhuma tabela foi encontrada no PDF. Verifique se o arquivo está correto."
        else:
            # Criar um escritor Excel para salvar múltiplas abas
            with pd.ExcelWriter("module/temp/output.xlsx", engine="openpyxl") as writer:
                for i, table in enumerate(tables, start=1):
                    # Transformar a tabela em DataFrame
                    df = pd.DataFrame(table)

                    # Ajustar cabeçalhos, se necessário
                    if not df.empty:
                        df.columns = df.iloc[0]  # Definir a primeira linha como cabeçalho
                        df = df[1:].reset_index(drop=True)  # Remover a linha de cabeçalho duplicada
                        df.columns = [str(col).strip() for col in df.columns]  # Ajustar cabeçalhos

                        # Remover colunas e linhas completamente vazias
                        df = df.dropna(how="all", axis=0)  # Linhas vazias
                        df = df.dropna(how="all", axis=1)  # Colunas vazias

                        # Salvar a tabela em uma aba nomeada (ex: "Página_1", "Página_2", etc.)
                        df.to_excel(writer, sheet_name=f"Tabela_Página_{i}", index=False)
                    else:
                        print(f"A tabela na página {i} está vazia. Pulando.")

            return "Tabelas temporárias exportadas para 'temp/output.xlsx'."
    except Exception as e:
        return f"Ocorreu um erro ao processar o PDF: {e}"

