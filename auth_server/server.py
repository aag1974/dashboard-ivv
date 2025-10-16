import os
import json
from flask import Flask, redirect, request, session, url_for, render_template
from google_auth_oauthlib.flow import Flow
from google.oauth2 import id_token
from google.auth.transport import requests as google_requests

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "chave-secreta-temporaria")

# ======== CONFIGURAÇÕES ========
REDIRECT_URI = "https://dashboard-ivv.onrender.com/callback"
SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile"
]
ALLOWED_USERS_FILE = os.path.join(os.path.dirname(__file__), "allowed_users.json")

# Lê o conteúdo do client_secret.json diretamente da variável de ambiente
CLIENT_SECRET_DATA = json.loads(os.environ["GOOGLE_CLIENT_SECRET_JSON"])

# ======== ROTAS ========

@app.route("/")
def index():
    """Exibe a página inicial (index.html) sempre que o usuário acessa a raiz."""
    return render_template("index.html")


@app.route("/login")
def login():
    flow = Flow.from_client_config(
        CLIENT_SECRET_DATA,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )
    authorization_url, state = flow.authorization_url(prompt="consent")
    session["state"] = state
    return redirect(authorization_url)


@app.route("/callback")
def callback():
    state = session.get("state")

    flow = Flow.from_client_config(
        CLIENT_SECRET_DATA,
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI
    )

    authorization_response = request.url
    flow.fetch_token(authorization_response=authorization_response)

    credentials = flow.credentials
    request_session = google_requests.Request()
    id_info = id_token.verify_oauth2_token(
        credentials._id_token,
        request_session,
        CLIENT_SECRET_DATA["web"]["client_id"]
    )

    user_email = id_info["email"]

    # Verifica se o usuário está autorizado
    with open(ALLOWED_USERS_FILE, "r") as f:
        allowed_users = json.load(f)

    if user_email not in allowed_users["allowed_emails"]:
        return "Usuário não autorizado. Contate o administrador."

    session["email"] = user_email
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    """Página do dashboard, protegida por login."""
    if "email" not in session:
        return redirect(url_for("login"))
    return render_template("dashboard.html", user_email=session["email"])


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


# ======== EXECUÇÃO LOCAL ========
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
