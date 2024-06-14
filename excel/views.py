import xlsxwriter
import pandas as pd
from django.http import HttpResponse
from rest_framework.views import APIView


class DownloadExcel(APIView):
    def get(self, request, *args, **kwargs):
        data = {
            'Nome': ['João', 'Maria', 'José'],
            'Idade': [23, 25, 30],
            'Salário': [1000, 2000, 3000]
        }
        df = pd.DataFrame(data)

        initial_date = "2024-01-01"
        final_date = "2024-12-31"
        company = "Exemplo Empresa"
        user_email = "user@example.com"

        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="dataframe.xlsx"'

        with pd.ExcelWriter(response, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Sheet1')

            # Formatação do título
            title = f'Relatório de Histórico de Consultas: {initial_date} até {final_date}'
            worksheet.write('A1', title, workbook.add_format({'bold': True}))

            # Formatação das informações adicionais
            worksheet.write('A3', f"Empresa: {company}" if company is not None else "Empresa: Filtro não aplicado")
            worksheet.write('A4', f"Filtro Data inicial: {initial_date}")
            worksheet.write('A5', f"Filtro Data final: {final_date}")
            worksheet.write('A6', f"Usuário Download: {user_email}")
            worksheet.write('A7', "")

            # Ajuste das larguras das colunas e formatação do cabeçalho do DataFrame
            for idx, column in enumerate(df.columns):
                column_letter = xlsxwriter.utility.xl_col_to_name(idx)
                max_length = max(df[column].astype(str).map(len).max(), len(column))
                adjusted_width = (max_length + 2) * 1.2
                worksheet.set_column(f'{column_letter}:{column_letter}', adjusted_width)

            # Formatação em negrito para o cabeçalho do DataFrame
            header_format = workbook.add_format({'bold': True})
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(8, col_num, value, header_format)

            # Escrever os dados do DataFrame começando da linha 9
            for row_num, (index, row) in enumerate(df.iterrows(), start=9):
                for col_num, value in enumerate(row):
                    worksheet.write(row_num, col_num, value)

            # Esconder linhas de grade
            worksheet.hide_gridlines(2)

            # Formatação de colunas específicas
            currency_format = workbook.add_format({'num_format': '$#,##0.00'})
            columns_to_format = ['Salário']  # Ajuste conforme necessário

            for column in columns_to_format:
                col_idx = df.columns.get_loc(column)
                worksheet.set_column(col_idx, col_idx, None, currency_format)

        return response
