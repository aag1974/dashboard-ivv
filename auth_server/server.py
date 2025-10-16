from flask import Flask, redirect, url_for, session, render_template
from authlib.integrations.flask_client import OAuth
import os
import json

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "chave-super-secreta")

# Configuração do OAuth com Google
oauth = OAuth(app)
google = oauth.register(
    name='google',
    client_id=os.getenv("GOOGLE_CLIENT_ID"),
    client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
    access_token_url='https://accounts.google.com/o/oauth2/token',
    authorize_url='https://accounts.google.com/o/oauth2/auth',
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
    return google.authorize_redirect(redirect_uri)

@app.route('/authorize')
def authorize():
    token = google.authorize_access_token()
    resp = google.get('userinfo')
    user_info = resp.json()
    user_email = user_info.get('email')

    # se o usuário não estiver autorizado → mostra aguardando.html
    if user_email not in allowed_users:
        return render_template('aguardando.html', email=user_email)

    session['google_token'] = token
    session['user'] = user_info
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
def dashboard():
    if 'google_token' not in session:
        return redirect(url_for('login'))
    user = session.get('user', {})
    return render_template('dashboard.html', user=user)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
