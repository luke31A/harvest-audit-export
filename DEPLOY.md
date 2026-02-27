# Deployment Guide — Harvest Audit Export

## One-time setup (do this in order)

### 1. Create a GitHub repository
Push this project to a **private** GitHub repo.
Streamlit Community Cloud will deploy directly from it.

```
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_ORG/harvest-audit-export.git
git push -u origin main
```

---

### 2. Register the Harvest OAuth application
Someone with **Harvest admin access** needs to do this once.

1. Go to https://id.getharvest.com/oauth2/clients/new
2. Fill in:
   - **Name:** Commit Harvest Audit
   - **Redirect URL:** `https://commitharvestaudit.streamlit.app`
   - **Kind:** Web (server-side)
3. Click Create — you'll get a **Client ID** and **Client Secret**
4. Save both somewhere secure — you'll need them in step 4

---

### 3. Deploy to Streamlit Community Cloud
1. Go to https://share.streamlit.io and sign in with GitHub
2. Click **New app**
3. Select your repo, branch `main`, and set **Main file path** to `app.py`
4. Set the app URL to `commitharvestaudit`
5. Click **Deploy** — don't add secrets yet, just get it deployed

---

### 4. Add secrets in Streamlit Cloud
1. In the Streamlit Cloud dashboard, open your app → **Settings → Secrets**
2. Paste the following (replacing with your real values from step 2):

```toml
[harvest]
client_id     = "your_client_id_here"
client_secret = "your_client_secret_here"
redirect_uri  = "https://commitharvestaudit.streamlit.app"
```

3. Click **Save** — the app will restart automatically

---

### 5. Test the login
Open https://commitharvestaudit.streamlit.app, click **Login with Harvest**,
and confirm you're redirected to Harvest and back successfully.

---

## Local development

```bash
# Install dependencies
pip install -r requirements.txt

# Create local secrets file
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit secrets.toml and add your credentials
# Set redirect_uri to http://localhost:8501 for local dev
# (also update this temporarily in your Harvest OAuth app settings)

# Run the app
streamlit run app.py
```

---

## CLI / .exe build (standalone, no Streamlit)
```bash
build.bat
```
Produces `dist/HarvestExport.exe` — runs without Python or internet access to the app.
