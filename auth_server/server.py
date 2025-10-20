from flask import Flask, render_template, send_from_directory, session, redirect, url_for
from authlib.integrations.flask_client import OAuth
import os
import json
import traceback
import secrets
from datetime import timedelta

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
# ROTAS PRINCIPAIS
# ==========================

@app.route('/')
def index():
    if 'google_token' in session:
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
    user_email = None
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
            return render_template("acesso_negado.html"), 403

        # Cria sessão autenticada
        session["user"] = {"email": user_email}
        session.permanent = True
        print(f"🎉 Sessão criada para {user_email}")
        return redirect(url_for("dashboard"))

    except Exception as e:
        print(f"❌ ERRO EM /AUTHORIZE: {e}")
        traceback.print_exc()
        return f"Erro interno durante autorização: {e}", 500


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
    session.pop("user", None)
    print("👋 Usuário desconectado")
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
