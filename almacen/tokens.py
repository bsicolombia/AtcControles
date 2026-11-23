import secrets

CLAVE = "tokens_form"

def nuevo_token(request):
    token = secrets.token_urlsafe(16)
    tokens = request.session.get(CLAVE, [])
    tokens.append(token)
    request.session[CLAVE] = tokens[-30:]
    return token

def token_valido(request):
    t = request.POST.get("form_token", "")
    return bool(t) and t in request.session.get(CLAVE, [])

def consumir_token(request):
    """Gasta el token. True si era válido; False si ya se había usado o no existe."""
    t = request.POST.get("form_token", "")
    tokens = request.session.get(CLAVE, [])
    if t and t in tokens:
        tokens.remove(t)
        request.session[CLAVE] = tokens
        request.session.modified = True
        return True
    return False

def token_para_render(request):
    """Reutiliza el token si sigue vigente; si no, emite uno nuevo."""
    t = request.POST.get("form_token", "")
    if t and t in request.session.get(CLAVE, []):
        return t
    return nuevo_token(request)

def devolver_token(request):
    """Si el guardado falló, el token vuelve a quedar válido."""
    t = request.POST.get("form_token", "")
    if t:
        tokens = request.session.get(CLAVE, [])
        if t not in tokens:
            tokens.append(t)
            request.session[CLAVE] = tokens[-200:]
            request.session.modified = True