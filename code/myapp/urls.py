from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('', views.index, name='index'),
    path('download_my_pdf',views.download_pdf, name='download_pdf'),
]
