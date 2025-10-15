from flask import Flask, redirect, url_for, session, request, render_template
from google_auth_oauthlib.flow import Flow
from google.oauth2 import id_token
from google.auth.transport import requests
import os
import json

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecretkey")

# Configurações do OAuth
CLIENT_SECRETS_FILE = "client_secret.json"
SCOPES = ["openid", "https://www.googleapis.com/auth/userinfo.email", "https://www.googleapis.com/auth/userinfo.profile"]
REDIRECT_URI = "https://dashboard-ivv.onrender.com/callback"

ALLOWED_USERS_FILE = "allowed_users.json"

# Página inicial
@app.route("/")
def index():
    if "email" in session:
        return render_template("index.html", user_email=session["email"])
    return redirect(url_for("login"))

# Rota de login
@app.route("/login")
def login():
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    authorization_url, state = flow.authorization_url(prompt="consent")
    session["state"] = state
    return redirect(authorization_url)

# Callback do Google
@app.route("/callback")
def callback():
    try:
        flow = Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )

        flow.fetch_token(authorization_response=request.url)

        credentials = flow.credentials
        request_session = requests.Request()
        id_info = id_token.verify_oauth2_token(
            credentials.id_token, request_session, flow.client_config["client_id"]
        )

        email = id_info.get("email")

        # Verifica se o usuário é permitido
        with open(ALLOWED_USERS_FILE, "r") as f:
            allowed_users = json.load(f).get("allowed_users", [])

        if email not in allowed_users:
            return render_template("aguardando.html", email=email)

        session["email"] = email
        return redirect(url_for("index"))

    except Exception as e:
        return f"Erro durante login: {e}"

# Logout
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

# Rota de dashboard (caso queira acessar diretamente)
@app.route("/dashboard")
def dashboard():
    if "email" not in session:
        return redirect(url_for("login"))
    return render_template("index.html", user_email=session["email"])

if __name__ == "__main__":
    app.run(debug=True)
