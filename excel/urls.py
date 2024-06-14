from django.urls import path
from .views import DownloadExcel

urlpatterns = [
    # outras rotas aqui...
    path('download_excel/', DownloadExcel.as_view(), name='download_excel'),]