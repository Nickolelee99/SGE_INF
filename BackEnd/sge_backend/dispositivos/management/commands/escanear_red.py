from django.core.management.base import BaseCommand
from dispositivos.models import EscaneoRed
from dispositivos.views import escanear_red
from datetime import date


class Command(BaseCommand):
    help = "Escaneo programado de red"

    def handle(self, *args, **kwargs):

        self.stdout.write("Iniciando escaneo...")

        dispositivos = escanear_red()

        self.stdout.write(
            f"Equipos encontrados: {len(dispositivos)}"
        )


        for dispositivo in dispositivos:

            hostname = dispositivo['hostname']

            codigo = ''.join(
                filter(str.isdigit, hostname)
            )


            EscaneoRed.objects.create(
                fecha_lote=date.today(),
                ip=dispositivo['ip'],
                hostname=hostname,
                codigo_actual=codigo
            )

        self.stdout.write(
            self.style.SUCCESS(
                "Escaneo almacenado correctamente"
            )
        )