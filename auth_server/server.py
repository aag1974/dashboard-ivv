from flask import Flask, redirect, url_for, session, render_template, session
from authlib.integrations.flask_client import OAuth
import os
import json
import traceback

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "chave-super-secreta")

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

@app.route('/')
def index():
    if 'google_token' in session:
        return redirect(url_for('dashboard'))
    return render_template('index.html')

@app.route('/login')
def login():
    redirect_uri = url_for('authorize', _external=True)
    print("🔍 Redirect URI gerado:", redirect_uri, flush=True)
    return google.authorize_redirect(redirect_uri)

@app.route("/authorize")
def authorize():
    try:
        # código original de autenticação
        token = oauth.google.authorize_access_token()
        user_info = oauth.google.parse_id_token(token)
        session["user"] = user_info
        return redirect("/dashboard")
    except Exception as e:
        import traceback
        print("❌ ERRO EM /AUTHORIZE:", e)
        traceback.print_exc()
        return f"Erro interno: {str(e)}", 500
    return redirect(url_for('dashboard'))

@app.route("/dashboard")
def dashboard():
    try:
        user = session.get("user")
        print("📍 Entrando em /dashboard com user:", user)
        return render_template("dashboard.html", user=user)
    except Exception as e:
        print("❌ ERRO EM /dashboard:", e)
        traceback.print_exc()
        return f"Erro interno: {e}", 500
    
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
