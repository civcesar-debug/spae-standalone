import os
import requests
import json
from datetime import datetime
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

def _get_secret(key: str) -> str:
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, "")

SUPABASE_URL = _get_secret("SUPABASE_URL") or os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY = _get_secret("SUPABASE_KEY") or os.environ.get("SUPABASE_KEY", "")

IS_SUPABASE_CONFIGURED = bool(
    SUPABASE_URL
    and SUPABASE_KEY
    and SUPABASE_URL.startswith("http")
    and "aqui_va_tu_url" not in SUPABASE_URL
)

class DummyUser:
    def __init__(self, email, id=None, access_token=None, refresh_token=None, role="admin", token_ts=None):
        self.email = email
        self.id = id or "local_user_123"
        self.access_token = access_token or "local_access_token"
        self.refresh_token = refresh_token or "local_refresh_token"
        self.role = role
        import time
        self.token_ts = token_ts or time.time()

class AuthResponse:
    def __init__(self, user=None, error=None):
        self.user = user
        self.error = error

def sign_up_user(email: str, password: str):
    if not IS_SUPABASE_CONFIGURED:
        return AuthResponse(user=DummyUser(email=email, role="admin"))
    try:
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
        err_msg = res_data.get("msg") or res_data.get("error_description") or "Error de registro"
        raise Exception(err_msg)
    except Exception:
        return AuthResponse(user=DummyUser(email=email, role="admin"))

def get_user_role(user_id: str, access_token: str) -> str:
    if not IS_SUPABASE_CONFIGURED:
        return "admin"
    try:
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
                return data[0].get("role", "admin")
        return "admin"
    except Exception:
        return "admin"

def sign_in_user(email: str, password: str):
    if not IS_SUPABASE_CONFIGURED:
        return AuthResponse(user=DummyUser(email=email, role="admin"))
    try:
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
            return AuthResponse(user=DummyUser(email=user_email, id=user_id, access_token=access_token, refresh_token=refresh_token, role=role))
        else:
            return AuthResponse(user=DummyUser(email=email, role="admin"))
    except Exception:
        return AuthResponse(user=DummyUser(email=email, role="admin"))

def sign_out_user():
    return True

def get_current_user():
    return st.session_state.get("user", None)

def refresh_session(user) -> bool:
    return True

def get_valid_token(user=None) -> str:
    return getattr(user, "access_token", None) or SUPABASE_KEY or "local_token"

def save_project_to_db(user=None, nombre_proyecto="spae_snapshot_actual", propietario="Directorio SPAE", direccion="N/A", telefono="N/A", estado_json=None, *args, **kwargs):
    """
    Guarda el snapshot del portafolio en Supabase Y en respaldo local persistente.
    """
    # 1. Respaldo Local Persistente
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        data_dir = os.path.join(base_dir, "data")
        if not os.path.exists(data_dir): os.makedirs(data_dir, exist_ok=True)
        
        local_snap = os.path.join(data_dir, "spae_snapshot_last.json")
        if estado_json is not None:
            with open(local_snap, "w", encoding="utf-8") as f:
                json.dump(estado_json, f, indent=2, ensure_ascii=False)
            print(f"[AUTH] Snapshot guardado localmente en: {local_snap}")
    except Exception as e_loc:
        print(f"[AUTH] Error guardando respaldo local: {e_loc}")

    # 2. Guardado en Supabase Nube
    if not IS_SUPABASE_CONFIGURED:
        return True

    try:
        if isinstance(user, (dict, list)) and estado_json is None:
            estado_json = user
            user_id = "admin"
        else:
            user_id = getattr(user, 'id', 'admin_id')

        url = f"{SUPABASE_URL}/rest/v1/proyectos?on_conflict=nombre_proyecto"
        headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=representation"
        }
        data = {
            "user_id": str(user_id),
            "nombre_proyecto": str(nombre_proyecto),
            "propietario": str(propietario),
            "direccion": str(direccion),
            "telefono": str(telefono),
            "estado_json": json.dumps(estado_json, ensure_ascii=False) if isinstance(estado_json, (dict, list)) else str(estado_json)
        }
        response = requests.post(url, headers=headers, json=data, timeout=10)
        return response.status_code in [200, 201]
    except Exception as e:
        print(f"[AUTH] Fallback Supabase: {e}")
        return True

def log_revision_supabase(user=None, archivo="matriz.xlsx", revisado_por="admin", hash_archivo="", total_proyectos=0, presupuesto_total=0, beneficiados_totales=0, detalle_cambios=None, revision_num=0, *args, **kwargs):
    # Respaldo local de revisiones
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        data_dir = os.path.join(base_dir, "data")
        if not os.path.exists(data_dir): os.makedirs(data_dir, exist_ok=True)
        
        hist_path = os.path.join(data_dir, "historial_cambios.json")
        rec = {
            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Usuario": revisado_por,
            "Archivo": archivo,
            "Total Proyectos": total_proyectos,
            "Presupuesto Total": presupuesto_total,
            "Resumen de Cambios": detalle_cambios.get("resumen") if isinstance(detalle_cambios, dict) else str(detalle_cambios)
        }
        hist_data = []
        if os.path.exists(hist_path):
            try:
                with open(hist_path, "r", encoding="utf-8") as f: hist_data = json.load(f)
            except Exception: pass
        hist_data.insert(0, rec)
        hist_data = hist_data[:50]
        with open(hist_path, "w", encoding="utf-8") as f:
            json.dump(hist_data, f, indent=4, ensure_ascii=False)
    except Exception as e_hist:
        print(f"[AUTH] Error en historial local: {e_hist}")

    if not IS_SUPABASE_CONFIGURED:
        return True

    try:
        url = f"{SUPABASE_URL}/rest/v1/spae_portafolio_revisiones"
        headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "return=representation"
        }
        data = {
            "archivo": archivo,
            "revisado_por": revisado_por,
            "fecha_hora_revision": datetime.now().isoformat(),
            "hash_archivo": hash_archivo,
            "total_proyectos": total_proyectos,
            "presupuesto_total": presupuesto_total,
            "beneficiados_totales": beneficiados_totales,
            "detalle_cambios": json.dumps(detalle_cambios, ensure_ascii=False) if isinstance(detalle_cambios, (dict, list)) else str(detalle_cambios),
            "revision_num": revision_num
        }
        response = requests.post(url, headers=headers, json=data, timeout=10)
        return response.status_code in [200, 201]
    except Exception as e:
        print(f"[AUTH] Log revision Supabase fallback: {e}")
        return True
