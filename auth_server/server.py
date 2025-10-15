from flask import Flask, redirect, url_for, session, request, render_template
from google_auth_oauthlib.flow import Flow
from google.oauth2 import id_token
from google.auth.transport import requests
import os
import json

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "dev_secret_key")  # altere no Render depois

# Configurações do OAuth (vindas das variáveis de ambiente)
GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://dashboard-ivv.onrender.com/callback")
ALLOWED_USERS_FILE = "auth_server/allowed_users.json"

# Função para criar o fluxo OAuth com os dados diretamente do ambiente
def create_flow():
    return Flow.from_client_config(
        {
            "web": {
                "client_id": GOOGLE_CLIENT_ID,
                "project_id": "dashboardivvlogin",
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": GOOGLE_CLIENT_SECRET,
                "redirect_uris": [REDIRECT_URI],
            }
        },
        scopes=[
            "https://www.googleapis.com/auth/userinfo.email",
            "https://www.googleapis.com/auth/userinfo.profile",
            "openid",
        ],
    )

# Rota inicial — botão de login
@app.route("/")
def index():
    if "google_id" in session:
        return render_template("dashboard.html", user_email=session["email"])
    return render_template("index.html")

# Rota de login (inicia o fluxo OAuth)
@app.route("/login")
def login():
    flow = create_flow()
    flow.redirect_uri = REDIRECT_URI
    authorization_url, state = flow.authorization_url(
        access_type="offline", include_granted_scopes="true"
    )
    session["state"] = state
    return redirect(authorization_url)

# Callback do Google (onde o usuário volta após login)
@app.route("/callback")
def callback():
    state = session["state"]
    flow = create_flow()
    flow.redirect_uri = REDIRECT_URI

    flow.fetch_token(authorization_response=request.url)
    credentials = flow.credentials

    request_session = requests.Request()
    user_info = id_token.verify_oauth2_token(
        credentials._id_token, request_session, GOOGLE_CLIENT_ID
    )

    email = user_info.get("email")

    # Carrega os usuários permitidos
    with open(ALLOWED_USERS_FILE, "r") as f:
        allowed_users = json.load(f)["allowed_users"]

    if email not in allowed_users:
        return render_template("aguardando.html", email=email)

    # Armazena sessão
    session["google_id"] = user_info.get("sub")
    session["email"] = email

    return redirect(url_for("dashboard"))

# Página principal após login
@app.route("/dashboard")
def dashboard():
    if "google_id" not in session:
        return redirect(url_for("index"))
    return render_template("dashboard.html", user_email=session["email"])

# Logout
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

# Servidor Flask
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
