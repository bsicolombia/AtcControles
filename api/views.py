from django.shortcuts import render

# Create your views here.
from django.db import transaction
from rest_framework import serializers, status
from rest_framework.decorators import action
from rest_framework.parsers import MultiPartParser
from rest_framework.response import Response
from rest_framework.viewsets import ReadOnlyModelViewSet

from almacen import services
from almacen.models import Asignacion
from operaciones.models import Operacion
from .serializers import OperacionSerializer, EjecutarSerializer, EvidenciaSerializer


class MisOperacionesViewSet(ReadOnlyModelViewSet):
    serializer_class = OperacionSerializer

    def get_queryset(self):
        tecnico = getattr(self.request.user, "perfil_tecnico", None)
        if not tecnico:
            return Operacion.objects.none()
        # El técnico SOLO ve lo suyo
        return Operacion.objects.filter(tecnico=tecnico).order_by("-fecha_operacion")

    @action(detail=True, methods=["post"])
    def ejecutar(self, request, pk=None):
        op = self.get_object()
        if op.estado != Operacion.Estado.ASIGNADA:
            raise serializers.ValidationError("La operación no está pendiente de ejecutar.")

        ser = EjecutarSerializer(data=request.data)
        ser.is_valid(raise_exception=True)

        try:
            with transaction.atomic():
                for m in ser.validated_data["materiales"]:
                    a = Asignacion.objects.get(pk=m["asignacion_id"], tecnico=op.tecnico)
                    services.instalar(
                        asignacion=a, usuario=request.user,
                        cantidad=m.get("cantidad"),
                        fecha_instalacion=ser.validated_data["fecha"],
                        operacion=op,
                        observaciones=ser.validated_data.get("observaciones", ""),
                    )
                op.estado = Operacion.Estado.EJECUTADA
                op.fecha_ejecucion = ser.validated_data["fecha"]
                op.save()
        except Asignacion.DoesNotExist:
            raise serializers.ValidationError("Material no asignado a este técnico.")
        except services.ErrorInventario as e:
            raise serializers.ValidationError(str(e))

        return Response({"estado": op.estado})

    @action(detail=True, methods=["post"], parser_classes=[MultiPartParser])
    def evidencias(self, request, pk=None):
        op = self.get_object()
        ser = EvidenciaSerializer(data=request.data)
        ser.is_valid(raise_exception=True)
        ser.save(operacion=op, usuario=request.user)
        return Response(ser.data, status=status.HTTP_201_CREATED)