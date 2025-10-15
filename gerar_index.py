# gerar_index.py
# Gera automaticamente um index.html com login Google + validação via Flask
# Autor: Alexandre Garcia (adaptado)
# Data: 2025-10

import os

CLIENT_ID = "314768988299-60uhkckmv0kqh0iah8l3tir7bdeemeci.apps.googleusercontent.com"

html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Dashboard Seguro</title>
  <script src="https://accounts.google.com/gsi/client" async defer></script>
  <script src="https://cdn.jsdelivr.net/npm/jwt-decode/build/jwt-decode.min.js"></script>
  <style>
    body {{
      font-family: Arial, sans-serif;
      text-align: center;
      margin-top: 50px;
      background-color: #f5f5f5;
      color: #333;
    }}
    #dashboard {{
      display: none;
      margin-top: 40px;
    }}
    button {{
      padding: 8px 15px;
      background: #4285F4;
      border: none;
      color: #fff;
      border-radius: 5px;
      cursor: pointer;
      margin-top: 15px;
    }}
    button:hover {{
      background: #3367D6;
    }}
    .g_id_signin {{
      display: inline-block;
      margin-top: 20px;
    }}
  </style>
</head>
<body>
  <h2>Entrar com Google</h2>

  <div id="g_id_onload"
       data-client_id="{CLIENT_ID}"
       data-callback="onSignIn"
       data-auto_prompt="false">
  </div>
  <div class="g_id_signin" data-type="standard"></div>

  <div id="dashboard">
    <h1>Bem-vindo ao Dashboard Seguro</h1>
    <p id="email"></p>
    <button onclick="logout()">Sair</button>
  </div>

  <script>
    async function onSignIn(response) {{
      const token = response.credential;
      const res = await fetch("/validate", {{
        method: "POST",
        headers: {{ "Content-Type": "application/json" }},
        body: JSON.stringify({{ token }})
      }});

      const data = await res.json();
      if (data.success) {{
        document.getElementById("email").innerText = "Usuário: " + data.email;
        document.querySelector(".g_id_signin").style.display = "none";
        document.getElementById("dashboard").style.display = "block";
      }} else {{
        alert(data.error);
      }}
    }}

    function logout() {{
      document.getElementById("dashboard").style.display = "none";
      document.querySelector(".g_id_signin").style.display = "block";
    }}
  </script>
</body>
</html>
"""

# salva o HTML
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("✅ Arquivo 'index.html' gerado com sucesso!")
print("👉 Agora execute: python3 server.py e acesse http://localhost:5000")
