import pandas as pd
import tabula


def process_pdf(path: str, pages: str):
    try:
        # Read tables from all specified pages
        tables = tabula.read_pdf(path, pages=pages, lattice=True, multiple_tables=True)
        if not tables or len(tables) == 0:
            return "Não encontramos tabelas no PDF. Por favor verifiique seo arquivo está correto."

        print(len(tables))
        for table in tables:
            # Criar um escritor Excel para salvar múltiplas abas
            with pd.ExcelWriter("module/temp/output.xlsx", engine="openpyxl") as writer:
                for i, tablein in enumerate(tables, start=1):
                    # Transformar a tabela em DataFrame
                    df = pd.DataFrame(tablein)

                    # Ajustar cabeçalhos, se necessário
                    if not df.empty:
                        # Checar se a primeira linha pode ser um cabeçalho válido
                        if all(isinstance(x, str) for x in df.iloc[0]) and not df.iloc[0].isnull().all():
                            # Usar a primeira linha como cabeçalho
                            df.columns = [str(col).strip() for col in df.iloc[0]]
                            df = df[1:].reset_index(drop=True)  # Remover a linha usada como cabeçalho
                        else:
                            # Criar cabeçalhos genéricos para remoção
                            df.columns = [f"remove" for i in range(1, len(df.columns) + 1)]

                        # Remover linhas ou colunas completamente vazias
                        df = df.dropna(how="all", axis=0)  # Linhas vazias
                        df = df.dropna(how="all", axis=1)  # Colunas vazias

                        # Verificar se a primeira linha (cabeçalho) ou qualquer outra linha contém o texto "remove" e excluir
                        if (df.columns == "remove").all():
                            # Se o cabeçalho for "remove", redefinir o cabeçalho para a próxima linha e remover a linha original
                            df.columns = df.iloc[0]  # Usa a segunda linha como cabeçalho
                            df = df[1:].reset_index(drop=True)  # Remove a primeira linha (agora excluída)
                        else:
                            # Caso "remove" esteja em outras linhas, excluí-las
                            df = df[~df.apply(lambda row: (row == "remove").all(), axis=1)]

                        # Remover colunas com cabeçalhos inválidos
                        df = df.loc[:, ~df.columns.str.contains("nan|Unnamed")]


                        # Salvar a tabela em uma aba nomeada (ex: "Página_1", "Página_2", etc.)
                        df.to_excel(writer, sheet_name=f"Tabela_Página_{i}", index=False)
                    else:
                        print(f"A tabela na página {i} está vazia. Pulando.")



    except Exception as e:
        return f"An error occurred while processing the PDF: {e}"
