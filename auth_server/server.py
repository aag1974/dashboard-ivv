from flask import Flask, render_template, send_from_directory, session, redirect, url_for, request
from authlib.integrations.flask_client import OAuth
import os
import json
import traceback
import secrets
from datetime import timedelta
import uuid  # (ADDED) para gerar session_id único

# ==========================
# CONFIGURAÇÃO PRINCIPAL
# ==========================
template_dir = os.path.join(os.path.dirname(__file__), "templates")
app = Flask(__name__, template_folder=template_dir)
app.secret_key = os.getenv("SECRET_KEY", "chave-super-secreta")
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="None"
)
app.permanent_session_lifetime = timedelta(hours=1)

# ==========================
# CONFIGURAÇÃO DO OAUTH (GOOGLE)
# ==========================
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id=os.getenv("GOOGLE_CLIENT_ID"),
    client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
    server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
    client_kwargs={'scope': 'openid email profile'}
)

# ==========================
# CAMINHO DO ARQUIVO allowed_users.json
# ==========================
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
ALLOWED_USERS_PATH = os.path.join(BASE_DIR, 'allowed_users.json')

if not os.path.exists(ALLOWED_USERS_PATH):
    raise FileNotFoundError(f"❌ Arquivo 'allowed_users.json' não encontrado em: {ALLOWED_USERS_PATH}")

with open(ALLOWED_USERS_PATH, 'r') as f:
    allowed_users = json.load(f)

print(f"✅ allowed_users.json carregado com sucesso ({len(allowed_users)} registros)")
print(f"📂 Caminho absoluto: {ALLOWED_USERS_PATH}")

# ==========================
# CONTROLE DE SESSÃO ÚNICA (ADDED)
# ==========================
SESSIONS_FILE = os.path.join(BASE_DIR, "sessions.json")
_active_sessions = {}  # mapa: email -> session_id

def _load_sessions():
    global _active_sessions
    if os.path.exists(SESSIONS_FILE):
        try:
            with open(SESSIONS_FILE, "r") as f:
                _active_sessions = json.load(f)
        except Exception:
            _active_sessions = {}
    else:
        _active_sessions = {}

def _save_sessions():
    try:
        with open(SESSIONS_FILE, "w") as f:
            json.dump(_active_sessions, f)
    except Exception as e:
        print("⚠️ Falha ao salvar sessions.json:", e)

_load_sessions()

def _set_active_session(user_email: str, session_id: str):
    """Define/atualiza a sessão ativa de um usuário e persiste em disco."""
    _active_sessions[user_email] = session_id
    _save_sessions()

def _clear_active_session(user_email: str):
    """Remove a sessão ativa registrada para um usuário e persiste em disco."""
    if user_email in _active_sessions:
        _active_sessions.pop(user_email, None)
        _save_sessions()

def _is_current_session_active(user_email: str, current_session_id: str) -> bool:
    """Confere se a sessão atual do navegador é a mesma registrada como ativa."""
    return _active_sessions.get(user_email) == current_session_id

# ==========================
# MIDDLEWARE: ENFORCE SINGLE SESSION (ADDED)
# ==========================
@app.before_request
def _enforce_single_session():
    # Rotas públicas/estáticas que não exigem checagem
    public_paths = ("/", "/login", "/authorize", "/acesso_negado", "/ping", "/static")
    if request.path.startswith(public_paths):
        return

    # Para rotas protegidas, exigimos sessão válida
    user = session.get("user")
    if not user:
        return redirect(url_for("login"))

    user_email = user.get("email")
    current_session_id = session.get("session_id")

    # Se não houver session_id (ex: sessão antiga), força re-login
    if not user_email or not current_session_id:
        session.clear()
        return redirect(url_for("login"))

    # Se a sessão registrada como ativa for diferente, derruba esta sessão
    if not _is_current_session_active(user_email, current_session_id):
        print(f"🧱 Sessão inválida/sobrescrita detectada para {user_email}. Forçando login.")
        session.clear()
        return redirect(url_for("login"))

# ==========================
# ROTAS PRINCIPAIS
# ==========================

@app.route('/')
def index():
    # Ajuste: considera a chave "user" que já é usada pelo app
    if 'user' in session:
        return redirect(url_for('dashboard'))
    return render_template('index.html')


@app.route("/login")
def login():
    """Inicia o login via Google"""
    nonce = secrets.token_urlsafe(16)
    session["nonce"] = nonce
    redirect_uri = url_for("authorize", _external=True)
    print("🔍 Redirect URI gerado:", redirect_uri)
    return oauth.google.authorize_redirect(redirect_uri, nonce=nonce)


@app.route("/authorize")
def authorize():
    """Callback do OAuth após login no Google"""
    try:
        token = oauth.google.authorize_access_token()
        user_info = oauth.google.parse_id_token(token, nonce=session.get("nonce"))
        user_email = user_info.get("email")
        print(f"✅ Login bem-sucedido: {user_email}")

        # Recarrega a lista de usuários permitidos
        with open(ALLOWED_USERS_PATH, 'r') as f:
            allowed_users = json.load(f)

        if user_email not in allowed_users:
            print(f"⛔ Acesso negado para {user_email}")
            session.clear()
            return redirect(url_for("acesso_negado"))

        # -------- Sessão única: gera e registra session_id (ADDED) --------
        new_session_id = str(uuid.uuid4())
        session["user"] = {"email": user_email}
        session["session_id"] = new_session_id
        session.permanent = True

        # Substitui qualquer sessão anterior deste usuário
        _set_active_session(user_email, new_session_id)
        print(f"🎉 Sessão criada/atualizada para {user_email} (session_id {new_session_id})")
        # ------------------------------------------------------------------

        return redirect(url_for("dashboard"))

    except Exception as e:
        print(f"❌ ERRO EM /AUTHORIZE: {e}")
        traceback.print_exc()
        return f"Erro interno durante autorização: {e}", 500


@app.route("/acesso_negado")
def acesso_negado():
    """Página mostrada quando o e-mail não está autorizado"""
    print("🚫 Redirecionado para acesso_negado")
    return render_template("acesso_negado.html"), 200

@app.route("/dashboard")
def dashboard():
    """Página principal (protegida)"""
    try:
        if "user" not in session:
            print("🚫 Acesso negado — redirecionando para login")
            return redirect(url_for("login"))

        print(f"✅ Usuário autenticado: {session['user']['email']}")
        caminho_templates = os.path.join(app.root_path, "templates")
        return send_from_directory(caminho_templates, "dashboard.html")

    except Exception as e:
        print("❌ ERRO EM /dashboard:", e)
        traceback.print_exc()
        return f"Erro interno: {e}", 500


@app.route("/logout")
def logout():
    """Finaliza a sessão"""
    try:
        user_email = session.get("user", {}).get("email")
        if user_email:
            _clear_active_session(user_email)  # (ADDED) limpa a sessão ativa registrada
        session.clear()
        print("👋 Usuário desconectado e sessão ativa removida")
    except Exception as e:
        print("⚠️ Erro durante logout:", e)
    return redirect("/")


@app.after_request
def add_header(response):
    """Evita cache em navegadores"""
    response.headers["Cache-Control"] = "no-store"
    return response


@app.route("/ping")
def ping():
    """Verificação de disponibilidade"""
    return "pong", 200


# ==========================
# EXECUÇÃO LOCAL
# ==========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
