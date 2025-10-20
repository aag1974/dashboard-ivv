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
    try:
        token = oauth.google.authorize_access_token()
        user_info = oauth.google.parse_id_token(token, nonce=session.get("nonce"))
        email = user_info.get("email", "").lower()

        # 🔒 Verifica se o e-mail está autorizado
        if email not in allowed_users:
            print(f"🚫 Acesso negado para: {email}")
            return render_template("acesso_negado.html", email=email), 403

        # Sessão válida e permanente (1h)
        session.permanent = True
        session["user"] = user_info

        print(f"✅ Login bem-sucedido: {email}")
        return redirect("/dashboard")

    except Exception as e:
        print("❌ ERRO EM /authorize:", e)
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

# Rota de ping (para manter vivo)
@app.route("/ping")
def ping():
    return "pong", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
