import pandas as pd
from rest_framework.views import APIView
from django.http import HttpResponse


class DownloadExcel(APIView):
    def get(self, request, *args, **kwargs):
        data = {
            'coluna1': [1, 2, 3],
            'coluna2': ['A', 'B', 'C'],
        }
        df = pd.DataFrame(data)

        response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="dataframe.xlsx"'

        with pd.ExcelWriter(response, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        return response