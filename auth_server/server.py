from flask import Flask, render_template, send_from_directory, session, redirect, url_for, request
from authlib.integrations.flask_client import OAuth
import os
import json
import traceback
import secrets
from datetime import timedelta, datetime
import uuid

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
# CONFIGURAÇÃO DE USUÁRIOS E PERFIS (fonte única: user_profiles.json)
# ==========================
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
USER_PROFILES_PATH = os.path.join(BASE_DIR, 'user_profiles.json')

def load_user_config():
    """Carrega configuração completa de usuários e perfis"""
    try:
        if not os.path.exists(USER_PROFILES_PATH):
            print(f"⚠️ user_profiles.json não encontrado em: {USER_PROFILES_PATH}")
            return {}
            
        with open(USER_PROFILES_PATH, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        users = config.get('users', {})
        active_users = [email for email, data in users.items() if data.get('active', False)]
        
        print(f"✅ user_profiles.json carregado com sucesso")
        print(f"📊 Total de usuários: {len(users)}")
        print(f"👥 Usuários ativos: {len(active_users)}")
        print(f"📂 Caminho: {USER_PROFILES_PATH}")
        
        return config
        
    except Exception as e:
        print(f"❌ Erro ao carregar user_profiles.json: {e}")
        return {}

def is_user_authorized(email):
    """Verifica se usuário está autorizado e ativo"""
    config = load_user_config()
    users = config.get('users', {})
    user_data = users.get(email, {})
    return user_data.get('active', False)

def get_user_profile(user_email):
    """Determina o perfil do usuário e atualiza último acesso"""
    try:
        config = load_user_config()
        users = config.get('users', {})
        user_data = users.get(user_email, {})
        
        if not user_data.get('active', False):
            print(f"⚠️ Usuário {user_email} está desativado")
            return 'viewer'
        
        profile = user_data.get('profile', 'viewer')
        print(f"📋 {user_email} → perfil: {profile}")
        
        # Atualizar último acesso
        update_user_last_access(user_email)
        
        return profile
        
    except Exception as e:
        print(f"❌ Erro ao obter perfil do usuário: {e}")
        return 'viewer'

def update_user_last_access(user_email):
    """Atualiza o timestamp de último acesso do usuário"""
    try:
        with open(USER_PROFILES_PATH, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        if user_email in config.get('users', {}):
            config['users'][user_email]['last_access'] = datetime.now().isoformat()
            
            with open(USER_PROFILES_PATH, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
    except Exception as e:
        print(f"⚠️ Erro ao atualizar último acesso: {e}")

# Carregar configuração inicial
user_config = load_user_config()
    
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

        # Verificar se usuário está autorizado e ativo
        if not is_user_authorized(user_email):
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
    """Página principal - serve dashboard baseado no perfil do usuário"""
    try:
        if "user" not in session:
            print("🚫 Acesso negado — redirecionando para login")
            return redirect(url_for("login"))

        user_email = session['user']['email']
        user_profile = get_user_profile(user_email)
        
        print(f"✅ Usuário autenticado: {user_email} (perfil: {user_profile})")
        
        # Determina qual dashboard servir baseado no perfil
        dashboard_file = f"dashboard_{user_profile}.html"
        
        # Verifica se arquivo específico do perfil existe
        caminho_templates = os.path.join(app.root_path, "templates")
        dashboard_path = os.path.join(caminho_templates, dashboard_file)
        
        if not os.path.exists(dashboard_path):
            print(f"⚠️ {dashboard_file} não encontrado, usando dashboard.html padrão")
            dashboard_file = "dashboard.html"
        
        print(f"📊 Servindo: {dashboard_file}")
        return send_from_directory(caminho_templates, dashboard_file)

    except Exception as e:
        print("❌ ERRO EM /dashboard:", e)
        traceback.print_exc()
        return f"Erro interno: {e}", 500

@app.route("/debug/profile")
def debug_profile():
    """Rota de debug para ver perfil do usuário (remover em produção)"""
    if "user" not in session:
        return "Não logado", 401
    
    user_email = session['user']['email']
    user_profile = get_user_profile(user_email)
    
    return f"""
    <h2>Debug - Informações do Usuário</h2>
    <p><strong>Email:</strong> {user_email}</p>
    <p><strong>Perfil:</strong> {user_profile}</p>
    <p><strong>Dashboard esperado:</strong> dashboard_{user_profile}.html</p>
    <br>
    <a href="/dashboard">Ir para Dashboard</a> | 
    <a href="/logout">Logout</a>
    """

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
