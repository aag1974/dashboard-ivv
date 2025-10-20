from flask import Flask, render_template, send_from_directory, session, redirect, url_for
from authlib.integrations.flask_client import OAuth
import os
import json
import traceback
import secrets
from datetime import timedelta

# Configuração principal
app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "chave-super-secreta")
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE="None"
)
app.permanent_session_lifetime = timedelta(hours=1)

# Configuração do OAuth (Google)
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id=os.getenv("GOOGLE_CLIENT_ID"),
    client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
    server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
    client_kwargs={'scope': 'openid email profile'}
)

# Carrega lista de e-mails permitidos
with open(os.path.join(os.path.dirname(__file__), 'allowed_users.json')) as f:
    allowed_users = json.load(f)

# Página inicial
@app.route('/')
def index():
    if 'google_token' in session:
        return redirect(url_for('dashboard'))
    return render_template('index.html')

# Login com Google
@app.route("/login")
def login():
    nonce = secrets.token_urlsafe(16)
    session["nonce"] = nonce
    redirect_uri = url_for("authorize", _external=True)
    print("🔍 Redirect URI gerado:", redirect_uri)
    return oauth.google.authorize_redirect(redirect_uri, nonce=nonce)

# Callback de autorização
@app.route("/authorize")
def authorize():
    user_email = None  # ✅ Garante que a variável exista mesmo se o try falhar
    try:
        token = oauth.google.authorize_access_token()
        user_info = oauth.google.parse_id_token(token, nonce=session.get("nonce"))
        user_email = user_info.get("email")
        print(f"✅ Login bem-sucedido: {user_email}")

        # Verifica se o e-mail está autorizado
        with open("allowed_users.json", "r") as f:
            allowed_users = json.load(f)

        if user_email not in allowed_users:
            print(f"⛔ Acesso negado para {user_email}")
            return render_template("acesso_negado.html"), 403

        # Autenticação bem-sucedida → cria sessão
        session["user"] = {"email": user_email}
        session.permanent = True
        return redirect(url_for("dashboard"))

    except Exception as e:
        print(f"❌ ERRO EM /AUTHORIZE: {e}")
        traceback.print_exc()
        return f"Erro interno durante autorização: {e}", 500

# Página do dashboard (protegida)
@app.route("/dashboard")
def dashboard():
    try:
        if "user" not in session:
            print("🚫 Acesso negado — redirecionando para login")
            return redirect(url_for("login"))

        print("✅ Usuário autenticado, servindo dashboard.html")
        caminho_templates = os.path.join(app.root_path, "templates")
        return send_from_directory(caminho_templates, "dashboard.html")
    except Exception as e:
        print("❌ ERRO EM /dashboard:", e)
        traceback.print_exc()
        return f"Erro interno: {e}", 500

# Logout
@app.route("/logout")
def logout():
    session.pop("user", None)
    print("👋 Usuário desconectado")
    return redirect("/")

@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-store"
    return response

# Rota de ping (para manter vivo)
@app.route("/ping")
def ping():
    return "pong", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
