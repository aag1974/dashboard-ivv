from flask import Flask, redirect, request, url_for, session, render_template
from google_auth_oauthlib.flow import Flow
from google.oauth2 import id_token
from google.auth.transport import requests
import pathlib, os, json
from email.mime.text import MIMEText
import smtplib

# ===== Config =====
app = Flask(__name__)
app.secret_key = "chave_super_secreta_segura"

# Em DEV permitimos http (localhost). Em PRODUÇÃO, NÃO defina isso.
if os.getenv("FLASK_ENV", "development") == "development":
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

BASE_URL = os.getenv("BASE_URL", "http://127.0.0.1:5000")  # em prod: https://seu-dominio
CLIENT_SECRETS_FILE = os.path.join(pathlib.Path(__file__).parent, "client_secret.json")
GOOGLE_CLIENT_ID = "314768988299-60uhkckmv0kqh0iah8l3tir7bdeemeci.apps.googleusercontent.com"

# E-mail do admin (opcional p/ notificação)
ADMIN_EMAIL = "alexopiniao@gmail.com"
ADMIN_APP_PASSWORD = os.getenv("ADMIN_APP_PASSWORD", "")  # senha de app do Gmail (opcional)

def load_allowed():
    path = os.path.join(pathlib.Path(__file__).parent, "allowed_users.json")
    if not os.path.exists(path): return []
    with open(path, "r") as f: return json.load(f)

def save_allowed(lst):
    path = os.path.join(pathlib.Path(__file__).parent, "allowed_users.json")
    with open(path, "w") as f: json.dump(lst, f, indent=4)

def notify_request(email):
    if not ADMIN_APP_PASSWORD:  # notificação é opcional
        return
    body = (
        f"O usuário {email} solicitou acesso ao Dashboard IVV.\n\n"
        f"Aprovar: {BASE_URL}/autorizar?email={email}\n"
    )
    msg = MIMEText(body)
    msg["Subject"] = "Solicitação de acesso - Dashboard IVV"
    msg["From"] = ADMIN_EMAIL
    msg["To"] = ADMIN_EMAIL
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls()
        s.login(ADMIN_EMAIL, ADMIN_APP_PASSWORD)
        s.send_message(msg)

def build_flow():
    return Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=[
            "openid",
            "https://www.googleapis.com/auth/userinfo.email",
            "https://www.googleapis.com/auth/userinfo.profile",
        ],
        redirect_uri=f"{BASE_URL}/callback",
    )

# ===== Rotas =====
@app.route("/")
def index():
    if "email" in session and session.get("authorized"):
        return redirect(url_for("dashboard"))
    return render_template("index.html")

@app.route("/login")
def login():
    flow = build_flow()
    authorization_url, state = flow.authorization_url(prompt="consent")
    session["state"] = state
    return redirect(authorization_url)

@app.route("/callback")
def callback():
    state = session.get("state")
    flow = build_flow()
    flow.fetch_token(authorization_response=request.url)

    creds = flow.credentials
    req = requests.Request()
    info = id_token.verify_oauth2_token(creds._id_token, req, GOOGLE_CLIENT_ID)

    email = info.get("email")
    session["email"] = email

    allowed = set(load_allowed())
    if email in allowed:
        session["authorized"] = True
        return redirect(url_for("dashboard"))
    else:
        session["authorized"] = False
        notify_request(email)
        return render_template("aguardando.html", email=email)

@app.route("/autorizar")
def autorizar():
    email = request.args.get("email", "").strip().lower()
    if not email:
        return "E-mail inválido."
    allowed = set(load_allowed())
    if email not in allowed:
        allowed.add(email)
        save_allowed(sorted(list(allowed)))
        msg = f"Acesso autorizado para {email}."
    else:
        msg = f"{email} já estava autorizado."
    return f"<h3>{msg}</h3><a href='{BASE_URL}'>Voltar</a>"

@app.route("/dashboard")
def dashboard():
    if not session.get("authorized"):
        return redirect(url_for("index"))
    return render_template("dashboard.html", user_email=session.get("email"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    # Em dev: python server.py  |  Em prod, use Gunicorn (ver passo 3)
    app.run(debug=(os.getenv("FLASK_ENV", "development")=="development"))
