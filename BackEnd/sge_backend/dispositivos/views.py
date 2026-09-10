import socket
import os
import math
import pandas as pd

from concurrent.futures import ThreadPoolExecutor, as_completed

from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_GET
from django.conf import settings

from .models import Inventario, EscaneoRed
from django.db.models import Max


import json
import re

from django.views.decorators.csrf import csrf_exempt

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image,
    KeepTogether,
    
)

# -----------------------------
# VISTAS BÁSICAS
# -----------------------------

def menu_principal(request):
    return render(request, 'dispositivos/menu_principal.html')



# -----------------------------
# INFORME TÉCNICO
# -----------------------------

# -----------------------------
# RESPONSABLES DEL INFORME
# -----------------------------

RESPONSABLES = {
    "Veronica Vanesa Garcia Vasquez":"ESPECIALISTA PROVINCIAL DE GESTIÓN DE SERVICIOS",
    "Jose Manuel Lainez Aguilar":"ESPECIALISTA PROVINCIAL DE APOYO DE AUDITORIA 1",
    "Nickole Cristina Lee Sagbay":"ASISTENTE PROVINCIAL DE APOYO DE AUDITORIA",
    "Alex Alonso Sandoval Pinzon":"ESPECIALISTA PROVINCIAL DE GESTION DE SERVICIOS",
}



def crear_informe_tecnico(request):
    return render(
        request,
        'dispositivos/informe_tecnico.html',
        {
            'responsables': RESPONSABLES,
        }
    )

def obtener_ruta_inventario():
    """
    Obtiene la ruta del archivo Excel del inventario.
    """
    return os.path.join(
        settings.BASE_DIR.parent,
        'datos',
        'inventario.xlsx'
    )


def limpiar_valor(valor):
    """
    Convierte valores vacíos o NaN de Excel en texto vacío.
    """
    if valor is None:
        return ''

    if isinstance(valor, float) and math.isnan(valor):
        return ''

    if pd.isna(valor):
        return ''

    return str(valor).strip()


def normalizar_codigo(valor):
    """Normaliza códigos provenientes de Excel."""
    valor = limpiar_valor(valor)
    if valor.endswith('.0') and valor[:-2].isdigit():
        valor = valor[:-2]
    return valor

# ============================================================
# BÚSQUEDA DE EQUIPOS POR PERSONA
# ============================================================

def leer_inventario_excel():
    """
    Lee el inventario tecnológico desde la hoja Hoja1.
    """

    ruta_excel = obtener_ruta_inventario()

    if not os.path.exists(ruta_excel):
        raise FileNotFoundError(
            f'No se encontró el archivo de inventario: {ruta_excel}'
        )

    df = pd.read_excel(
        ruta_excel,
        sheet_name='Hoja1',
        dtype=str
    )

    # Limpiar nombres de columnas
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # Reemplazar valores vacíos / NaN
    df = df.fillna('')

    # Limpiar espacios
    for columna in df.columns:
        df[columna] = (
            df[columna]
            .astype(str)
            .str.strip()
        )

    return df


@require_GET
def obtener_personas(request):
    """
    Devuelve las personas que tienen equipos
    registrados en el inventario.
    """

    try:

        df = leer_inventario_excel()

        columna_persona = 'APELLIDOS Y NOMBRES COMPLETOS'

        if columna_persona not in df.columns:

            return JsonResponse({
                'encontrado': False,
                'mensaje': (
                    f'No existe la columna '
                    f'"{columna_persona}" en el inventario.'
                )
            }, status=500)

        personas = (
            df[columna_persona]
            .loc[
                df[columna_persona].str.strip() != ''
            ]
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        return JsonResponse({
            'encontrado': True,
            'personas': personas
        })

    except Exception as e:

        return JsonResponse({
            'encontrado': False,
            'mensaje': str(e)
        }, status=500)


@require_GET
def obtener_equipos_persona(request, nombre_persona):
    """
    Devuelve todos los equipos asignados
    a una persona.
    """

    try:

        nombre_persona = str(nombre_persona).strip()

        df = leer_inventario_excel()

        columna_persona = (
            'APELLIDOS Y NOMBRES COMPLETOS'
        )

        resultado = df[
            df[columna_persona].str.upper()
            ==
            nombre_persona.upper()
        ]

        if resultado.empty:

            return JsonResponse({
                'encontrado': False,
                'mensaje': (
                    f'No se encontraron equipos '
                    f'asignados a {nombre_persona}.'
                )
            }, status=404)


        equipos = []


        for _, activo in resultado.iterrows():

            equipo = {

                'codigo_actual': normalizar_codigo(
                    activo.get('CÓDIGO DEL BIEN')
                ),

                'descripcion': limpiar_valor(
                    activo.get(
                        'DESCRIPCIÓN /CARACTERÍSTICAS DEL BIEN'
                    )
                ),

                'marca': limpiar_valor(
                    activo.get('MARCA')
                ),

                'modelo': limpiar_valor(
                    activo.get('MODELO')
                ),

                'serie': limpiar_valor(
                    activo.get('NÚMERO DE SERIE')
                ),

                'estado': limpiar_valor(
                    activo.get('ESTADO')
                ),

                'procesador': limpiar_valor(
                    activo.get('PROCESADOR/CAPACIDAD')
                ),

                'memoria': limpiar_valor(
                    activo.get('MEMORIA GB')
                ),

                'disco': limpiar_valor(
                    activo.get('DISCO GB')
                ),

                'sistema_operativo': limpiar_valor(
                    activo.get('SISTEMA OPERATIVO')
                ),

                'persona': limpiar_valor(
                    activo.get(
                        'APELLIDOS Y NOMBRES COMPLETOS'
                    )
                ),

                'ubicacion': limpiar_valor(
                    activo.get(
                        'UBICACION FISICA EDIFICIO'
                    )
                ),
            }


            if equipo['codigo_actual']:

                equipos.append(equipo)


        return JsonResponse({

            'encontrado': True,

            'persona': nombre_persona,

            'total': len(equipos),

            'equipos': equipos

        })


    except Exception as e:

        return JsonResponse({

            'encontrado': False,

            'mensaje': str(e)

        }, status=500)

@require_GET

def buscar_activo(request, codigo_activo):
    """Busca un activo en la hoja Hoja1 del inventario tecnológico 2026."""
    try:
        codigo_busqueda = normalizar_codigo(codigo_activo)
        ruta_excel = obtener_ruta_inventario()

        if not os.path.exists(ruta_excel):
            return JsonResponse({
                'encontrado': False,
                'mensaje': f'No se encontró el archivo de inventario: {ruta_excel}'
            }, status=500)

        df = pd.read_excel(ruta_excel, sheet_name='Hoja1', dtype=str)
        df.columns = df.columns.astype(str).str.strip()

        columna_codigo = 'CÓDIGO DEL BIEN'
        if columna_codigo not in df.columns:
            return JsonResponse({
                'encontrado': False,
                'mensaje': f'No se encontró la columna "{columna_codigo}" en el inventario.'
            }, status=500)

        df['_codigo_normalizado'] = (
            df[columna_codigo].fillna('').astype(str).map(normalizar_codigo)
        )
        resultado = df[df['_codigo_normalizado'] == codigo_busqueda]

        if resultado.empty:
            return JsonResponse({
                'encontrado': False,
                'mensaje': f'No se encontró ningún activo con el código {codigo_busqueda}.'
            }, status=404)

        activo = resultado.iloc[0]

        return JsonResponse({
            'encontrado': True,
            'codigo_actual': normalizar_codigo(activo.get('CÓDIGO DEL BIEN')),
            'codigo_bien': normalizar_codigo(activo.get('CÓDIGO DEL BIEN')),
            'tipo_bien': limpiar_valor(activo.get('TIPO DE BIEN')),
            'descripcion': limpiar_valor(activo.get('DESCRIPCIÓN /CARACTERÍSTICAS DEL BIEN')),
            'marca': limpiar_valor(activo.get('MARCA')),
            'modelo': limpiar_valor(activo.get('MODELO')),
            'serie': limpiar_valor(activo.get('NÚMERO DE SERIE')),
            'procesador': limpiar_valor(activo.get('PROCESADOR/CAPACIDAD')),
            'memoria': limpiar_valor(activo.get('MEMORIA GB')),
            'disco': limpiar_valor(activo.get('DISCO GB')),
            'sistema_operativo': limpiar_valor(activo.get('SISTEMA OPERATIVO')),
            'estado': limpiar_valor(activo.get('ESTADO')),
            'piso': limpiar_valor(activo.get('PISO')),
            'anterior_ubicacion': limpiar_valor(activo.get('UBICACION FISICA EDIFICIO')),
            'anterior_custodio': limpiar_valor(activo.get('APELLIDOS Y NOMBRES COMPLETOS')),
            'cedula_custodio': limpiar_valor(activo.get('CÉDULA DE USUARIO FINAL')),
            'cargo': limpiar_valor(activo.get('CARGO')),
            'unidad_administrativa': limpiar_valor(activo.get('UNIDAD ADMINISTRATIVA')),
            'color': limpiar_valor(activo.get('COLOR')),
            'material': limpiar_valor(activo.get('MATERIAL')),
            'fecha_compra': limpiar_valor(activo.get('FECHA DE COMPRA')),
            'proveedor': limpiar_valor(activo.get('PROVEEDOR')),
            'periodo_garantia': limpiar_valor(activo.get('PERIODO DE GARANTIA')),
            'fecha_expiracion_garantia': limpiar_valor(activo.get('FECHA DE EXPIRACIÓN DE GARANTÍA')),
            'observaciones_inventario': limpiar_valor(activo.get('OBSERVACIONES')),
        })

    except Exception as e:
        print(f'Error al buscar activo {codigo_activo}: {e}')
        return JsonResponse({
            'encontrado': False,
            'mensaje': str(e)
        }, status=500)


def texto_pdf(valor):
    """
    Convierte valores vacíos en un espacio visible para el PDF.
    """
    if valor is None:
        return ''

    return str(valor).strip()


def escapar_pdf(valor):
    """
    Evita problemas con caracteres especiales dentro de Paragraph.
    """
    valor = texto_pdf(valor)

    return (
        valor
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
    )


def nombre_archivo_seguro(nombre):
    """
    Genera un nombre de archivo seguro.
    """
    nombre = texto_pdf(nombre)

    if not nombre:
        return 'informe_tecnico'

    nombre = re.sub(
        r'[^A-Za-z0-9_-]+',
        '_',
        nombre
    )

    return nombre


def dibujar_pie_pagina(canvas, doc):
    """
    Pie de página del informe.
    """
    canvas.saveState()

    canvas.setFont('Helvetica', 8)

    canvas.drawCentredString(
        A4[0] / 2,
        1.2 * cm,
        f'Página {doc.page}'
    )

    canvas.restoreState()


ruta_logo = os.path.join(
    settings.BASE_DIR,
    'staticfiles',
    'logo-contraloria.png'
)


@csrf_exempt
def generar_informe_pdf(request):
    

    if request.method != 'POST':
        return JsonResponse(
            {
                'error': 'Método no permitido.'
            },
            status=405
        )

    try:

        datos = json.loads(request.body.decode('utf-8'))

        numero_informe = texto_pdf(
            datos.get('numero_informe')
        )

        fecha = texto_pdf(
            datos.get('fecha')
        )

        ciudad = texto_pdf(
            datos.get('ciudad', 'Guayaquil')
        )

        anterior_custodio = texto_pdf(
            datos.get('anterior_custodio')
        )

        anterior_ubicacion = texto_pdf(
            datos.get('anterior_ubicacion')
        )

        nuevo_custodio = texto_pdf(
            datos.get('nuevo_custodio')
        )

        nueva_ubicacion = texto_pdf(
            datos.get('nueva_ubicacion')
        )

        observaciones = texto_pdf(
            datos.get('observaciones')
        )

        elaborado_por = texto_pdf(
            datos.get('elaborado_por')
        )

        cargo_elaborado = texto_pdf(
            datos.get('cargo_elaborado')
        )

        supervisado_por = texto_pdf(
            datos.get('supervisado_por')
        )

        cargo_supervisado = texto_pdf(
            datos.get('cargo_supervisado')
        )

        equipos = datos.get('equipos', [])


        # Crear respuesta PDF
        respuesta = HttpResponse(
            content_type='application/pdf'
        )

        nombre_archivo = nombre_archivo_seguro(
            numero_informe
        )

        respuesta[
            'Content-Disposition'
        ] = (
            f'attachment; '
            f'filename="{nombre_archivo}.pdf"'
        )


        documento = SimpleDocTemplate(
            respuesta,
            pagesize=A4,
            rightMargin=1.5 * cm,
            leftMargin=1.8 * cm,
            topMargin=1.5 * cm,
            bottomMargin=2 * cm,
        )


        estilos = getSampleStyleSheet()


        estilo_titulo = ParagraphStyle(
            'TituloInforme',
            parent=estilos['Heading1'],
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            fontSize=15,
            leading=18,
            spaceAfter=8,
        )


        estilo_subtitulo = ParagraphStyle(
            'SubtituloInforme',
            parent=estilos['Normal'],
            alignment=TA_CENTER,
            fontName='Helvetica',
            fontSize=10,
            leading=13,
            spaceAfter=15,
        )


        estilo_seccion = ParagraphStyle(
            'SeccionInforme',
            parent=estilos['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=10,
            textColor=colors.white,
            backColor=colors.HexColor('#2c3e50'),
            borderPadding=6,
            spaceBefore=10,
            spaceAfter=8,
        )


        estilo_normal = ParagraphStyle(
            'NormalInforme',
            parent=estilos['Normal'],
            fontName='Helvetica',
            fontSize=8.5,
            leading=11,
        )


        estilo_celda = ParagraphStyle(
            'CeldaInforme',
            parent=estilo_normal,
            fontSize=7.5,
            leading=9,
        )


        estilo_celda_negrita = ParagraphStyle(
            'CeldaNegrita',
            parent=estilo_celda,
            fontName='Helvetica-Bold',
        )


        elementos = []

        # =========================================
        # ENCABEZADO CON LOGO
        # =========================================

        estilo_titulo_header = ParagraphStyle(
            'TituloHeader',
            parent=estilos['Normal'],
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            fontSize=15,
            leading=18,
            spaceAfter=6,
        )

        estilo_numero_header = ParagraphStyle(
            'NumeroHeader',
            parent=estilos['Normal'],
            alignment=TA_CENTER,
            fontName='Helvetica',
            fontSize=10,
            leading=13,
        )


        # Logo
        logo = Image(
            ruta_logo,
            width=3.0 * cm,
            height=3.0 * cm
        )

        # Texto del encabezado
        texto_header = [
            Paragraph(
                'CONTRALORÍA GENERAL DEL ESTADO<br/>'
                'Dirección Provincial de Guayas<br/>'
                'Unidad Administrativa<br/>',
                estilo_titulo_header
            ),

            Paragraph(
                (
                    '<b>INFORME TÉCNICO DE BIENES TECNOLÓGICOS N°:</b> '
                    f'{escapar_pdf(numero_informe)}'
                ),
                estilo_numero_header
            ),
        ]


        # Tabla invisible para mantener logo y texto
        tabla_header = Table(
            [
                [
                    logo,
                    texto_header
                ]
            ],
            colWidths=[
                4.0 * cm,
                13.4 * cm,
            ],
            rowHeights=[
                3.2 * cm
            ]
        )


        tabla_header.setStyle(
            TableStyle([
                # Sin bordes
                (
                    'BOX',
                    (0, 0),
                    (-1, -1),
                    0,
                    colors.white
                ),

                (
                    'INNERGRID',
                    (0, 0),
                    (-1, -1),
                    0,
                    colors.white
                ),

                # Alineación vertical
                (
                    'VALIGN',
                    (0, 0),
                    (-1, -1),
                    'MIDDLE'
                ),

                # Logo a la izquierda
                (
                    'ALIGN',
                    (0, 0),
                    (0, 0),
                    'LEFT'
                ),

                # Texto centrado
                (
                    'ALIGN',
                    (1, 0),
                    (1, 0),
                    'CENTER'
                ),

                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    0
                ),

                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    0
                ),
            ])
        )


        elementos.append(tabla_header)

        elementos.append(
            Spacer(1, 0.3 * cm)
        )



        tabla_fecha = Table(
            [
                [
                    Paragraph(
                        '<b>Ciudad</b>',
                        estilo_celda_negrita
                    ),
                    Paragraph(
                        escapar_pdf(ciudad),
                        estilo_celda
                    ),

                    Paragraph(
                        '<b>Fecha</b>',
                        estilo_celda_negrita
                    ),
                    Paragraph(
                        escapar_pdf(fecha),
                        estilo_celda
                    ),
                ]
            ],
            colWidths=[
                2.2 * cm,
                5.5 * cm,
                2 * cm,
                5.5 * cm,
            ]
        )

        tabla_fecha.setStyle(
            TableStyle([
                (
                    'GRID',
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),
                (
                    'VALIGN',
                    (0, 0),
                    (-1, -1),
                    'MIDDLE'
                ),
                (
                    'BACKGROUND',
                    (0, 0),
                    (0, 0),
                    colors.HexColor('#eeeeee')
                ),
                (
                    'BACKGROUND',
                    (2, 0),
                    (2, 0),
                    colors.HexColor('#eeeeee')
                ),
                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    5
                ),
                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    6
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    6
                ),
            ])
        )

        elementos.append(tabla_fecha)

        elementos.append(Spacer(1, 0.6 * cm)) 
        # =========================================
        # CAMBIO DE CUSTODIA
        # =========================================

        tabla_cambio = Table(
            [
                 [
                    Paragraph('<b>CAMBIO DE CUSTODIO Y UBICACIÓN</b>', estilo_celda_negrita), 
                    ''
                ],


                [
                    Paragraph(
                        '<b>INFORMACIÓN ANTERIOR</b>',
                        estilo_celda_negrita
                    ),
                    Paragraph(
                        '<b>NUEVA INFORMACIÓN</b>',
                        estilo_celda_negrita
                    ),
                ],
                [
                    Paragraph(
                        (
                            '<b>Custodio:</b><br/>'
                            f'{escapar_pdf(anterior_custodio)}'
                            '<br/><br/>'
                            '<b>Ubicación:</b><br/>'
                            f'{escapar_pdf(anterior_ubicacion)}'
                        ),
                        estilo_normal
                    ),
                    Paragraph(
                        (
                            '<b>Custodio:</b><br/>'
                            f'{escapar_pdf(nuevo_custodio)}'
                            '<br/><br/>'
                            '<b>Ubicación:</b><br/>'
                            f'{escapar_pdf(nueva_ubicacion)}'
                        ),
                        estilo_normal
                    ),
                ]
            ],
            colWidths=[
                8.5 * cm,
                8.5 * cm,
            ],
            hAlign='LEFT'
        )

        # 2. Aplicamos el TableStyle para fusionar y dar formato
        tabla_cambio.setStyle(TableStyle([
            # Fusión: Une columna 0 a columna 1 en la fila 0
            ('SPAN', (0, 0), (1, 0)),
            
            # Alineación centrada para el título principal (fila 0)
            ('ALIGN', (0, 0), (1, 0), 'LEFT'),
            
            # Fondo decorativo para diferenciar el título principal (opcional)
            ('BACKGROUND', (0, 0), (1, 0), colors.lightgrey),
            
            # Configuración de bordes para toda la tabla
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        tabla_cambio.setStyle(
            TableStyle([
                (
                    'GRID',
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.grey
                ),
                (
                    'BACKGROUND',
                    (0, 0),
                    (-1, 0),
                    colors.HexColor('#eeeeee')
                ),
                (
                    'VALIGN',
                    (0, 0),
                    (-1, -1),
                    'TOP'
                ),
                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
            ])
        )

        elementos.append(tabla_cambio)

        elementos.append(Spacer(1, 0.6 * cm))



        # =========================================
        # DETALLE DE EQUIPOS
        # =========================================
        elementos.append(
            Paragraph('<b>DETALLE DE LOS EQUIPOS</b>', estilo_normal)
        )

        encabezados_equipos = [
            'N.º',
            'Código',
            'Tipo',
            'Marca',
            'Modelo',
            'N.º de serie',
            'Estado',
        ]
        filas_equipos = [
            [
                Paragraph(f'<b>{encabezado}</b>', estilo_celda_negrita)
                for encabezado in encabezados_equipos
            ]
        ]

        for indice, equipo in enumerate(equipos, start=1):
            filas_equipos.append([
                Paragraph(str(indice), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('codigo')), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('descripcion')), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('marca')), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('modelo')), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('serie')), estilo_celda),
                Paragraph(escapar_pdf(equipo.get('estado')), estilo_celda),
            ])

        if len(filas_equipos) == 1:
            filas_equipos.append([
                Paragraph('Sin equipos agregados', estilo_celda),
                '',
                '',
                '',
                '',
                '',
                '',
            ])

        # Se agrega hAlign='LEFT' para alinear la tabla a la izquierda
        tabla_equipos = Table(
            filas_equipos,
            colWidths=[
                0.7 * cm,
                2.4 * cm,
                3.5 * cm,
                2.4 * cm,
                2.7 * cm,
                3 * cm,
                2.3 * cm,
            ],
            repeatRows=1,
            hAlign='LEFT',
        )

        tabla_equipos.setStyle(
            TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eeeeee')),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ])
        )

        elementos.append(tabla_equipos)
        elementos.append(Spacer(1, 0.3 * cm))



        # =========================================
        # PERIFÉRICOS
        # =========================================

        perifericos = []

        for indice, equipo in enumerate(
            equipos,
            start=1
        ):

            accesorios = []

            if equipo.get('teclado'):
                accesorios.append('Teclado')

            if equipo.get('mouse'):
                accesorios.append('Mouse')

            if accesorios:

                perifericos.append(
                    [
                        Paragraph(
                            f'<b>Equipo {indice}</b>',
                            estilo_celda
                        ),

                        Paragraph(
                            escapar_pdf(
                                ', '.join(accesorios)
                            ),
                            estilo_celda
                        ),
                    ]
                )


        if perifericos:


            elementos.append(
                Paragraph(
                    '<b>PERIFÉRICOS INCLUIDOS</b>',
                    estilo_normal
                )
            )


            tabla_perifericos = Table(
                perifericos,
                colWidths=[
                    3 * cm,
                    14 * cm,
                ]
            )
            tabla_perifericos.hAlign = 'LEFT'

            tabla_perifericos.setStyle(
                TableStyle([
                    (
                        'GRID',
                        (0, 0),
                        (-1, -1),
                        0.5,
                        colors.grey
                    ),
                    (
                        'BACKGROUND',
                        (0, 0),
                        (0, -1),
                        colors.HexColor('#eeeeee')
                    ),
                    (
                        'LEFTPADDING',
                        (0, 0),
                        (-1, -1),
                        5
                    ),
                    (
                        'RIGHTPADDING',
                        (0, 0),
                        (-1, -1),
                        5
                    ),
                    (
                        'TOPPADDING',
                        (0, 0),
                        (-1, -1),
                        5
                    ),
                    (
                        'BOTTOMPADDING',
                        (0, 0),
                        (-1, -1),
                        5
                    ),
                ])
            )

            elementos.append(tabla_perifericos)
            elementos.append(Spacer(1, 0.6 * cm)) 

    
        
        # =========================================
        # ESPECIFICACIONES TÉCNICAS
        # =========================================

        equipos_con_especificaciones = []

        for indice, equipo in enumerate(equipos, start=1):

            especificaciones = []

            campos_especificaciones = [
                ('Procesador', equipo.get('procesador')),
                ('Memoria RAM', equipo.get('memoria')),
                ('Disco', equipo.get('disco')),
                ('Sistema operativo', equipo.get('sistema_operativo')),
            ]

            # Procesar las especificaciones de ESTE equipo
            for nombre, valor in campos_especificaciones:

                if texto_pdf(valor):

                    especificaciones.append(
                        [
                            Paragraph(
                                f'<b>{nombre}</b>',
                                estilo_celda_negrita
                            ),
                            Paragraph(
                                escapar_pdf(valor),
                                estilo_celda
                            ),
                        ]
                    )

            # Guardar las especificaciones de ESTE equipo
            if especificaciones:

                equipos_con_especificaciones.append(
                    {
                        'numero': indice,
                        'codigo': equipo.get('codigo'),
                        'descripcion': equipo.get('descripcion'),
                        'datos': especificaciones,
                    }
                )


        if equipos_con_especificaciones:

            elementos.append(
                Paragraph(
                    '<b>ESPECIFICACIONES TÉCNICAS</b>',
                    estilo_normal
                )
            )


            for item in equipos_con_especificaciones:

                texto_titulo_equipo = (
                    f'<b>Equipo {item["numero"]}</b> - '
                    f'{escapar_pdf(item["descripcion"])} '
                    f'({escapar_pdf(item["codigo"])})'
                )

                filas_tabla_actual = [
                    [
                        Paragraph(
                            texto_titulo_equipo,
                            estilo_celda_negrita
                        ),
                        ''
                    ]
                ] + item['datos']

                tabla_especificaciones = Table(
                    filas_tabla_actual,
                    colWidths=[
                        4 * cm,
                        13 * cm,
                    ]
                )

                tabla_especificaciones.hAlign = 'LEFT'

                tabla_especificaciones.setStyle(
                    TableStyle([
                        (
                            'SPAN',
                            (0, 0),
                            (1, 0)
                        ),
                        (
                            'BACKGROUND',
                            (0, 0),
                            (1, 0),
                            colors.HexColor('#dddddd')
                        ),
                        (
                            'TOPPADDING',
                            (0, 0),
                            (1, 0),
                            6
                        ),
                        (
                            'BOTTOMPADDING',
                            (0, 0),
                            (1, 0),
                            6
                        ),
                        (
                            'BACKGROUND',
                            (0, 1),
                            (0, -1),
                            colors.HexColor('#eeeeee')
                        ),
                        (
                            'GRID',
                            (0, 0),
                            (-1, -1),
                            0.5,
                            colors.grey
                        ),
                        (
                            'VALIGN',
                            (0, 0),
                            (-1, -1),
                            'MIDDLE'
                        ),
                        (
                            'LEFTPADDING',
                            (0, 0),
                            (-1, -1),
                            5
                        ),
                        (
                            'RIGHTPADDING',
                            (0, 0),
                            (-1, -1),
                            5
                        ),
                        (
                            'TOPPADDING',
                            (0, 1),
                            (-1, -1),
                            5
                        ),
                        (
                            'BOTTOMPADDING',
                            (0, 1),
                            (-1, -1),
                            5
                        ),
                    ])
                )

                elementos.append(
                    KeepTogether([
                        tabla_especificaciones,
                        Spacer(1, 0.5 * cm)
                    ])
                )


        # =========================================
        # OBSERVACIONES
        # =========================================

        elementos.append(
            Paragraph(
                '<b>OBSERVACIONES</b>',
                estilo_normal
            )
        )

        texto_observaciones = escapar_pdf(
            observaciones
        ).replace('\n', '<br/>')

        tabla_observaciones = Table(
            [
                [
                    Paragraph(
                        texto_observaciones,
                        estilo_normal
                    )
                ]
            ],
            colWidths=[
                17 * cm
            ],
            style=[
                (
                    'BOX',
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.grey
                ),
                (
                    'MINROWHEIGHT',
                    (0, 0),
                    (-1, -1),
                    3 * cm
                ),
                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
            ]
        )

        # Alinear la tabla con el margen izquierdo
        tabla_observaciones.hAlign = 'LEFT'

        elementos.append(tabla_observaciones)

        elementos.append(
            Spacer(1, 0.5 * cm)
        )

                
        # =========================================
        # RESPONSABLES Y FIRMAS
        # =========================================

        elementos.append(
            Paragraph(
                '<b>RESPONSABLES</b>',
                estilo_normal
            )
        )


        tabla_firmas = Table(
            [
                [
                    Paragraph(
                        (
                            '<br/><br/><br/><br/>'
                            '<b>____________________________</b>'
                            '<br/>'
                            '<b>ELABORADO POR</b>'
                            '<br/><br/>'
                            f'{escapar_pdf(elaborado_por)}'
                            '<br/>'
                            f'{escapar_pdf(cargo_elaborado)}'
                        ),
                        estilo_normal
                    ),

                    Paragraph(
                        (
                            '<br/><br/><br/><br/>'
                            '<b>____________________________</b>'
                            '<br/>'
                            '<b>SUPERVISADO POR</b>'
                            '<br/><br/>'
                            f'{escapar_pdf(supervisado_por)}'
                            '<br/>'
                            f'{escapar_pdf(cargo_supervisado)}'
                        ),
                        estilo_normal
                    ),
                ]
            ],
            colWidths=[
                8.5 * cm,
                8.5 * cm,
            ]
        )
        tabla_firmas.hAlign = 'LEFT'

        tabla_firmas.setStyle(
            TableStyle([
                (
                    'ALIGN',
                    (0, 0),
                    (-1, -1),
                    'CENTER'
                ),
                (
                    'VALIGN',
                    (0, 0),
                    (-1, -1),
                    'TOP'
                ),
                (
                    'BOX',
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),
                (
                    'INNERGRID',
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),
                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    10
                ),
            ])
        )

        elementos.append(tabla_firmas)


        # Construir PDF
        documento.build(
            elementos,
            onFirstPage=dibujar_pie_pagina,
            onLaterPages=dibujar_pie_pagina,
        )


        return respuesta


    except Exception as error:

        print(
            f'Error al generar informe PDF: {error}'
        )

        return JsonResponse(
            {
                'error': str(error)
            },
            status=500
        )
# -----------------------------
# ESCANEO OPTIMIZADO CON HILOS
# -----------------------------

def probar_ip(ip):
    """
    Intenta conectarse a un host en el puerto 135 (NetBIOS)
    """
    try:
        socket.setdefaulttimeout(0.1)

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((ip, 135))
            hostname = socket.getfqdn(ip)

            return {
                'ip': ip,
                'hostname': hostname
            }

    except:
        return None


def escanear_red():
    ip_bases = ['172.16.18.', '172.16.19.']
    ips = [base + str(i) for base in ip_bases for i in range(1, 255)]

    dispositivos = []
    max_workers = 100  # ajustable según tu red

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(probar_ip, ip) for ip in ips]

        for future in as_completed(futures):
            resultado = future.result()
            if resultado:
                dispositivos.append(resultado)

    return dispositivos


# -----------------------------
# LECTURA DB + CORRELACIÓN
# -----------------------------


def leer_inventario(request):
    try:

        # ==========================
        # 1. Leer inventario
        # ==========================
        inventario = list(
            Inventario.objects.filter(
                descripcion__in=[
                    'COMPUTADOR DE ESCRITORIO',
                    'COMPUTADOR PORTATIL',
                    'IMPRESORA ALTO VOLUMEN'
                ]
            ).order_by(
                'unidad_administrativa',
                'ubicacion',
                'descripcion'
            ).values(
                'unidad_administrativa',
                'ubicacion',
                'descripcion',
                'codigo_actual',
                'nombre_usuario'
            )
        )

        # ==========================
        # 2. Obtener el último lote
        # ==========================

        ultima_fecha = EscaneoRed.objects.aggregate(
            Max('fecha_lote')
        )['fecha_lote__max']

        hostnames_por_activo = {}

        if ultima_fecha:

            ultimo_escaneo = EscaneoRed.objects.filter(
                fecha_lote=ultima_fecha
            )

            for equipo in ultimo_escaneo:

                hostnames_por_activo[equipo.codigo_actual] = {
                    'hostname': equipo.hostname,
                    'ip': equipo.ip
                }

        print(f"Último lote: {ultima_fecha}")
        print(f"Equipos conectados: {len(hostnames_por_activo)}")

        # ==========================
        # 3. Relacionar inventario
        # ==========================

        for item in inventario:

            codigo = str(item['codigo_actual']).strip()

            if codigo in hostnames_por_activo:

                item['hostname'] = hostnames_por_activo[codigo]['hostname']
                item['ip'] = hostnames_por_activo[codigo]['ip']
                item['conectado'] = True

            else:

                item['hostname'] = ''
                item['ip'] = ''
                item['conectado'] = False

        return render(
            request,
            'dispositivos/lista_dispositivos.html',
            {
                'dispositivos': inventario,
                'title': f'Inventario - Último escaneo {ultima_fecha}'
            }
        )

    except Exception as e:
        return HttpResponse(f"Error: {e}")