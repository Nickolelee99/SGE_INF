from django.db import models

class Inventario(models.Model):
    id = models.BigIntegerField(primary_key=True)

    unidad_administrativa = models.CharField(max_length=255, null=True, blank=True)
    ubicacion = models.CharField(max_length=255, null=True, blank=True)
    descripcion = models.TextField(null=True, blank=True)
    codigo_actual = models.CharField(max_length=50)
    nombre_usuario = models.TextField(null=True, blank=True)

    class Meta:
        db_table = 'inventario'
        managed = False



class EscaneoRed(models.Model):
    fecha_escaneo = models.DateTimeField(auto_now_add=True)
    ip = models.CharField(max_length=50)
    hostname = models.CharField(max_length=255)
    codigo_actual = models.CharField(max_length=50)
    fecha_lote = models.DateField()
    

    class Meta:
        db_table = 'escaneo_red'
        managed = False