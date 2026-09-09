from django.urls import path
from . import views

app_name = 'dispositivos'

urlpatterns = [
    path('', views.menu_principal, name='menu_principal'),

    path(
        'lista_dispositivos/',
        views.leer_inventario,
        name='lista_dispositivos'
    ),

    path(
        'informe-tecnico/',
        views.crear_informe_tecnico,
        name='crear_informe_tecnico'
    ),

    path(
        'api/activo/<str:codigo_activo>/',
        views.buscar_activo,
        name='buscar_activo'
    ),

    path(
        'generar-informe-pdf/',
        views.generar_informe_pdf,
        name='generar_informe_pdf'
    ),

    
]