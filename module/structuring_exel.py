import os
import pandas as pd
from openpyxl.styles import Alignment


def structuring(name_file: str):
    # Certifique-se de que o diretório "module/temp" existe
    os.makedirs("module/temp", exist_ok=True)

    # Verificar se o diretório contém arquivos
    if not os.listdir("module/temp"):
        return "O diretório 'module/temp' está vazio ou não existe."

    for file_name in os.listdir("module/temp"):
        # Verificar se o arquivo é do tipo Excel
        if file_name.endswith((".xlsx", ".xls")):
            file_path = os.path.join("module/temp", file_name)
            excel_file = pd.ExcelFile(file_path, engine="openpyxl")

            transformed_sheets = {}

            for sheet_name in excel_file.sheet_names:
                file_read = pd.read_excel(file_path, sheet_name=sheet_name)

                # Verificar se as colunas obrigatórias existem
                required_columns = ["LOCAL", "MODALIDADE\nDE VENDA"]
                if not all(col in file_read.columns for col in required_columns):
                    raise ValueError(
                        f"As colunas obrigatórias {required_columns} estão ausentes na aba {sheet_name} do arquivo {file_name}."
                    )

                # Transformar para formato longo
                melted_data = file_read.melt(
                    id_vars=required_columns,
                    var_name="Data",
                    value_name="Valor"
                )

                # Renomear colunas
                melted_data.rename(columns={
                    "LOCAL": "Local",
                    "MODALIDADE\nDE VENDA": "Modalidade de Venda",
                    "Data": "Data",
                    "Valor": "Valor"
                }, inplace=True)

                # Converter tipos
                melted_data["Local"] = melted_data["Local"].astype(str)
                melted_data["Modalidade de Venda"] = melted_data["Modalidade de Venda"].astype(str)
                melted_data["Data"] = pd.to_datetime(melted_data["Data"], format='%d.%m.%Y', errors='coerce')

                if melted_data["Data"].isna().any():
                    print(f"Aviso: Datas inválidas encontradas na aba {sheet_name} do arquivo {file_name}.")

                melted_data["Valor"] = (
                    melted_data["Valor"]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                    .pipe(pd.to_numeric, errors='coerce')
                )

                if melted_data["Valor"].isna().any():
                    print(f"Aviso: Valores inválidos na coluna 'Valor' na aba {sheet_name} do arquivo {file_name}.")

                transformed_sheets[sheet_name] = melted_data

            # Salvar o resultado
            os.makedirs("data", exist_ok=True)
            output_file = f"data/{name_file}.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for sheet_name, transformed_data in transformed_sheets.items():
                    transformed_data.to_excel(writer, index=False, sheet_name=sheet_name)
                    worksheet = writer.sheets[sheet_name]

                    for col in worksheet.columns:
                        if col[0].value == "Data":
                            for cell in col[1:]:
                                cell.number_format = "DD/MM/YYYY"
                                cell.alignment = Alignment(horizontal="center")
                        elif col[0].value == "Valor":
                            for cell in col[1:]:
                                cell.number_format = "#,##0.00"
                                cell.alignment = Alignment(horizontal="right")

            return f"\033[32mArquivo processado e salvo em: {output_file}\033[0m"

    return "Nenhum arquivo Excel encontrado em 'module/temp'."
