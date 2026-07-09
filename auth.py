import os
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# Soporte dual: Streamlit Cloud (st.secrets) y local (.env)
def _get_secret(key: str) -> str:
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, "")

SUPABASE_URL = _get_secret("SUPABASE_URL")
SUPABASE_KEY = _get_secret("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Faltan SUPABASE_URL o SUPABASE_KEY. Configuralos en .env (local) o en Secrets de Streamlit Cloud.")

import requests

class DummyUser:
    def __init__(self, email, id=None, access_token=None, refresh_token=None, role="free", token_ts=None):
        self.email = email
        self.id = id
        self.access_token = access_token
        self.refresh_token = refresh_token
        self.role = role
        import time
        self.token_ts = token_ts or time.time()  # Timestamp del último login/refresh

class AuthResponse:
    def __init__(self, user=None, error=None):
        self.user = user
        self.error = error

def sign_up_user(email: str, password: str):
    url = f"{SUPABASE_URL}/auth/v1/signup"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    data = {"email": email, "password": password}
    response = requests.post(url, headers=headers, json=data, timeout=5)
    res_data = response.json()
    if response.status_code == 200:
        user_id = res_data.get("user", {}).get("id")
        user_email = res_data.get("user", {}).get("email", email)
        access_token = res_data.get("access_token")
        refresh_token = res_data.get("refresh_token")
        return AuthResponse(user=DummyUser(email=user_email, id=user_id, access_token=access_token, refresh_token=refresh_token))
    err_msg = response.json().get("msg") or response.json().get("error_description") or "Error desconocido de registro"
    raise Exception(err_msg)

def get_user_role(user_id: str, access_token: str) -> str:
    url = f"{SUPABASE_URL}/rest/v1/profiles?id=eq.{user_id}&select=role"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {access_token or SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers, timeout=5)
    if response.status_code == 200:
        data = response.json()
        if data and len(data) > 0:
            return data[0].get("role", "free")
        else:
            # Fallback: Create profile if it doesn't exist (for old users)
            insert_url = f"{SUPABASE_URL}/rest/v1/profiles"
            insert_headers = {
                "apikey": SUPABASE_KEY,
                "Authorization": f"Bearer {access_token or SUPABASE_KEY}",
                "Content-Type": "application/json",
                "Prefer": "return=representation"
            }
            insert_data = {"id": user_id, "role": "free"}
            requests.post(insert_url, headers=insert_headers, json=insert_data, timeout=5)
            return "free"
    return "free"

def sign_in_user(email: str, password: str):
    url = f"{SUPABASE_URL}/auth/v1/token?grant_type=password"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    data = {"email": email, "password": password}
    response = requests.post(url, headers=headers, json=data, timeout=5)
    res_data = response.json()
    if response.status_code == 200:
        user_id = res_data.get("user", {}).get("id")
        user_email = res_data.get("user", {}).get("email", email)
        access_token = res_data.get("access_token")
        refresh_token = res_data.get("refresh_token")

        role = get_user_role(user_id, access_token)
        if email == "civcesar@gmail.com":
            role = "admin"

        return AuthResponse(user=DummyUser(email=user_email, id=user_id, access_token=access_token, refresh_token=refresh_token, role=role))
    err_msg = res_data.get("error_description") or res_data.get("msg") or "Credenciales invalidas"
    raise Exception(err_msg)

def sign_out_user():
    return True

def get_current_user():
    return st.session_state.get("user", None)

def refresh_session(user) -> bool:
    """Renueva el access_token usando el refresh_token de Supabase. Retorna True si tuvo éxito."""
    rt = getattr(user, "refresh_token", None)
    if not rt:
        return False
    url = f"{SUPABASE_URL}/auth/v1/token?grant_type=refresh_token"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    res = requests.post(url, headers=headers, json={"refresh_token": rt}, timeout=5)
    if res.status_code == 200:
        import time
        data = res.json()
        user.access_token = data.get("access_token", user.access_token)
        user.refresh_token = data.get("refresh_token", user.refresh_token)
        user.token_ts = time.time()
        # Actualizar session_state para que todos los módulos lo vean
        st.session_state["auth_user"] = user
        return True
    return False

def get_valid_token(user) -> str | None:
    """Devuelve el access_token vigente, refrescando automáticamente si está por vencer.
    Supabase emite JWTs de 1 h (3600 s por defecto);
    refrescamos a los 55 min (3300 s) para tener margen."""
    import time
    if not user:
        return None
    token_ts = getattr(user, "token_ts", 0)
    elapsed = time.time() - token_ts
    # Refrescar a los 3300 s (55 min) para sesiones de 1 h
    if elapsed > 3300:
        ok = refresh_session(user)
        if not ok:
            return None  # Sesión caducada — el módulo debe pedir re-login
    return getattr(user, "access_token", None)

def save_project_to_db(user, nombre_proyecto, propietario, direccion, telefono, estado_json):
    import json
    user_id = getattr(user, 'id', None)
    if not user_id:
        raise Exception("Usuario no autenticado")
    # Siempre usar un token válido (auto-refresca si está por expirar)
    access_token = get_valid_token(user) or SUPABASE_KEY
    url = f"{SUPABASE_URL}/rest/v1/proyectos?on_conflict=nombre_proyecto"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Prefer": "resolution=merge-duplicates,return=representation"
    }
    data = {
        "user_id": user_id,
        "nombre_proyecto": nombre_proyecto,
        "propietario": propietario,
        "direccion": direccion,
        "telefono": telefono,
        "estado_json": json.dumps(estado_json) if not isinstance(estado_json, str) else estado_json
    }
    response = requests.post(url, headers=headers, json=data, timeout=10)
    if response.status_code not in (200, 201):
        raise Exception(f"Error al guardar: {response.status_code} - {response.text}")
    return response.json()

def log_revision_supabase(user, archivo, revisado_por, hash_archivo, total_proyectos, presupuesto_total, beneficiados_totales, detalle_cambios, revision_num):
    import json
    import datetime
    user_id = getattr(user, 'id', None)
    if not user_id:
        raise Exception("Usuario no autenticado")
    access_token = get_valid_token(user) or SUPABASE_KEY
    url = f"{SUPABASE_URL}/rest/v1/spae_portafolio_revisiones"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    data = {
        "archivo": archivo,
        "revisado_por": revisado_por,
        "fecha_hora_revision": datetime.datetime.now().isoformat(),
        "hash_archivo": hash_archivo,
        "total_proyectos": total_proyectos,
        "presupuesto_total": presupuesto_total,
        "beneficiados_totales": beneficiados_totales,
        "detalle_cambios": json.dumps(detalle_cambios) if not isinstance(detalle_cambios, str) else detalle_cambios,
        "snapshot_version": revision_num
    }
    response = requests.post(url, headers=headers, json=data, timeout=10)
    if response.status_code not in (200, 201):
        raise Exception(f"Error al guardar log de revision: {response.status_code} - {response.text}")
    return response.json()


def get_projects_from_db(user):
    user_id = getattr(user, 'id', None)
    if not user_id:
        raise Exception("Usuario no autenticado")
    # Siempre usar un token válido (auto-refresca si está por expirar)
    access_token = get_valid_token(user) or SUPABASE_KEY
    url = f"{SUPABASE_URL}/rest/v1/proyectos?user_id=eq.{user_id}&order=created_at.desc"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers, timeout=10)
    if response.status_code != 200:
        raise Exception(f"Error al obtener proyectos: {response.status_code} - {response.text}")
    return response.json()

def delete_project_from_db(user, nombre_proyecto):
    user_id = getattr(user, 'id', None)
    # Siempre usar un token válido (auto-refresca si está por expirar)
    access_token = get_valid_token(user) or getattr(user, 'access_token', None)
    if not user_id:
        raise Exception("Usuario no autenticado")
    
    import requests
    # Importante: para DELETE se requiere la PK o filtro
    url = f"{SUPABASE_URL}/rest/v1/proyectos?user_id=eq.{user_id}&nombre_proyecto=eq.{nombre_proyecto}"
    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {access_token or SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    response = requests.delete(url, headers=headers, timeout=5)
    if response.status_code not in (200, 204):
        raise Exception(f"Error al eliminar: {response.status_code} - {response.text}")
    return True

def reset_password(email: str):
    url = f"{SUPABASE_URL}/auth/v1/recover"
    headers = {
        "apikey": SUPABASE_KEY,
        "Content-Type": "application/json"
    }
    data = {"email": email}
    response = requests.post(url, headers=headers, json=data, timeout=5)
    if response.status_code != 200:
        res_data = response.json()
        err_msg = res_data.get("msg") or res_data.get("error_description") or "Error al enviar correo de recuperación"
        raise Exception(err_msg)
    return True
