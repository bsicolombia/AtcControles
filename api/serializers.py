from rest_framework import serializers

from almacen.models import Asignacion, Producto
from operaciones.models import Operacion
from evidencias.models import EvidenciaInstalacion   # ajusta al nombre real de tu app


# ------------------------------------------------------------
# LECTURA: lo que el celular recibe
# ------------------------------------------------------------

class ProductoMiniSerializer(serializers.ModelSerializer):
    nombre_completo = serializers.CharField(read_only=True)
    es_serializado = serializers.BooleanField(read_only=True)

    class Meta:
        model = Producto
        fields = ["id", "codigo_producto", "nombre_completo", "es_serializado"]


class MaterialAsignadoSerializer(serializers.ModelSerializer):
    """Material que el técnico tiene en su poder y puede instalar."""

    producto = ProductoMiniSerializer(read_only=True)
    pendiente = serializers.IntegerField(source="cantidad_pendiente", read_only=True)
    unidad_codigo = serializers.CharField(
        source="unidad_serializada.codigo", read_only=True, default=None
    )
    unidad_serie = serializers.CharField(
        source="unidad_serializada.numero_serie", read_only=True, default=None
    )

    class Meta:
        model = Asignacion
        fields = [
            "id", "producto", "cantidad", "pendiente",
            "unidad_codigo", "unidad_serie", "estado",
        ]


class EvidenciaSerializer(serializers.ModelSerializer):
    """Sirve para leer y para crear (multipart desde el celular)."""

    class Meta:
        model = EvidenciaInstalacion
        fields = ["id", "imagen", "observacion", "client_uuid", "fecha_subida"]
        read_only_fields = ["id", "fecha_subida"]

    def validate_imagen(self, imagen):
        if imagen.size > 8 * 1024 * 1024:
            raise serializers.ValidationError("La imagen no puede superar 8 MB.")
        return imagen

    def validate_client_uuid(self, valor):
        # Reintento por mala señal: si ya existe, no se duplica.
        if valor and EvidenciaInstalacion.objects.filter(client_uuid=valor).exists():
            raise serializers.ValidationError("Esta evidencia ya fue enviada.")
        return valor


class OperacionSerializer(serializers.ModelSerializer):
    tipo_operacion_texto = serializers.CharField(
        source="get_tipo_operacion_display", read_only=True
    )
    estado_texto = serializers.CharField(source="get_estado_display", read_only=True)
    materiales_disponibles = serializers.SerializerMethodField()
    evidencias = EvidenciaSerializer(many=True, read_only=True)

    class Meta:
        model = Operacion
        fields = [
            "id", "numero_orden", "numero_cuenta_contrato",
            "tipo_operacion", "tipo_operacion_texto",
            "estado", "estado_texto",
            "fecha_operacion", "fecha_ejecucion", "observaciones",
            "materiales_disponibles", "evidencias",
        ]
        read_only_fields = fields   # el celular nunca edita la operación directamente

    def get_materiales_disponibles(self, op):
        """Lo que el técnico tiene pendiente de instalar (no depende de la operación)."""
        if not op.tecnico_id:
            return []
        pendientes = (
            Asignacion.objects
            .filter(tecnico_id=op.tecnico_id, estado=Asignacion.Estado.ASIGNADA)
            .select_related("producto", "unidad_serializada")
        )
        pendientes = [a for a in pendientes if a.cantidad_pendiente > 0]
        return MaterialAsignadoSerializer(pendientes, many=True).data


# ------------------------------------------------------------
# ESCRITURA: lo que el celular envía al ejecutar
# ------------------------------------------------------------

class MaterialInstaladoSerializer(serializers.Serializer):
    asignacion_id = serializers.IntegerField()
    cantidad = serializers.IntegerField(min_value=1, required=False)  # omitir si es unidad con código


class EjecutarSerializer(serializers.Serializer):
    fecha = serializers.DateField()
    observaciones = serializers.CharField(required=False, allow_blank=True, default="")
    materiales = MaterialInstaladoSerializer(many=True, allow_empty=False)

    def validate_materiales(self, items):
        ids = [m["asignacion_id"] for m in items]
        if len(ids) != len(set(ids)):
            raise serializers.ValidationError("Material repetido en la lista.")
        return items