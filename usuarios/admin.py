from django.contrib import admin


class PerfilTecnicoAdmin(admin.ModelAdmin): 
    list_display = ("usuario", "cedula", "supervisor") 
    search_fields = ( "usuario__username", "usuario__first_name", "usuario__last_name", "cedula", ) 
    list_filter = ("supervisor",)