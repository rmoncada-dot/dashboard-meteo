# 🌬️ Dashboard Meteo Stazioni — Pipeline Completa

## Struttura del progetto

```
meteo_pipeline/
│
├── excel/
│   └── macro_meteo.bas          # Macro VBA da incollare in Excel
│
├── streamlit/
│   ├── app.py                   # App Streamlit (dashboard interattiva)
│   └── requirements.txt         # Dipendenze Python
│
├── Dockerfile                   # Container per Google Cloud Run
├── .github/
│   └── workflows/
│       └── deploy.yml           # CI/CD automatico su push → Cloud Run
│
└── README.md                    # Questa guida
```

---

## 📊 STEP 1 — Excel con Macro VBA

### Setup
1. Apri Excel → crea un nuovo file → salvalo come **`.xlsm`** (Cartella di lavoro con macro)
2. Vai in **Strumenti → Macro → Visual Basic Editor** (oppure `Alt + F11`)
3. Nel menu VBA: **Insert → Module**
4. Copia e incolla tutto il contenuto di `excel/macro_meteo.bas`
5. Chiudi l'editor VBA

### Utilizzo
1. In Excel: **Strumenti → Macro → Macro → `ImportaCSVECreaaDashboard`** → Esegui
2. Seleziona la cartella con i file CSV della stazione
3. La macro creerà automaticamente:
   - **Foglio "Dati Raw"** — tutti i CSV uniti in una tabella
   - **Foglio "Statistiche"** — medie mensili, P50/P75/P90, shear α, disponibilità
   - **Foglio "Dashboard"** — 4 grafici + KPI cards
   - **File `export_streamlit.json`** — dati pronti per Streamlit

> ⚠️ Se Excel blocca le macro: **File → Opzioni → Centro protezione → Impostazioni macro → Abilita tutte le macro**

---

## 🐍 STEP 2 — App Streamlit (locale)

### Prerequisiti
- Python 3.10+
- pip

### Installazione e avvio locale
```bash
# Clona il repo
git clone https://github.com/TUONOME/NOMEREPO.git
cd NOMEREPO

# Crea ambiente virtuale
python -m venv .venv
source .venv/bin/activate        # Mac/Linux
# oppure: .venv\Scripts\activate  # Windows

# Installa dipendenze
pip install -r streamlit/requirements.txt

# Avvia l'app
streamlit run streamlit/app.py
```
L'app sarà disponibile su `http://localhost:8501`

### Sorgenti dati supportate
- **CSV diretti** — carica i file CSV mensili dalla sidebar
- **JSON da Excel** — carica `export_streamlit.json` generato dalla macro VBA
- **Demo** — dati precaricati della stazione G243043 Durrà

---

## 🔧 STEP 3 — Setup GitHub

### Crea il repository
1. Vai su [github.com](https://github.com) → **New repository**
2. Nome: `dashboard-meteo`
3. Visibilità: **Public** (o Private se preferisci)
4. Clicca **Create repository**

### Carica il codice
```bash
cd meteo_pipeline
git init
git add .
git commit -m "Prima versione dashboard meteo"
git branch -M main
git remote add origin https://github.com/TUONOME/dashboard-meteo.git
git push -u origin main
```

---

## ☁️ STEP 4 — Google Cloud Console

### 4a. Crea progetto GCP
1. Vai su [console.cloud.google.com](https://console.cloud.google.com)
2. Click sul menu a tendina in alto → **Nuovo progetto**
3. Nome: `dashboard-meteo` → **Crea**
4. Copia il **Project ID** (es. `dashboard-meteo-123456`)

### 4b. Abilita le API necessarie
Nella Cloud Console → **API e servizi → Libreria**, abilita:
- **Cloud Run Admin API**
- **Cloud Build API**
- **Container Registry API**

Oppure da terminale con gcloud:
```bash
gcloud services enable run.googleapis.com cloudbuild.googleapis.com containerregistry.googleapis.com
```

### 4c. Crea Service Account per GitHub Actions
1. **IAM e amministrazione → Account di servizio → Crea account di servizio**
2. Nome: `github-deployer`
3. Ruoli da assegnare:
   - `Cloud Run Admin`
   - `Cloud Build Editor`
   - `Storage Admin`
   - `Service Account User`
4. Crea chiave JSON: **Tasti → Aggiungi chiave → JSON** → scarica il file

### 4d. Aggiungi Secrets su GitHub
Nel tuo repo GitHub → **Settings → Secrets and variables → Actions → New repository secret**:

| Nome Secret      | Valore                                    |
|------------------|-------------------------------------------|
| `GCP_PROJECT_ID` | Il tuo Project ID (es. `dashboard-meteo-123456`) |
| `GCP_SA_KEY`     | Contenuto intero del file JSON scaricato  |

### 4e. Deploy automatico
Da questo momento, ogni `git push` sul branch `main` attiva automaticamente:
1. Build del container Docker
2. Push su Google Container Registry
3. Deploy su Cloud Run

Il workflow è in `.github/workflows/deploy.yml`.

### 4f. Verifica il deploy
Dopo il primo push, vai su **GitHub → Actions** per seguire il deploy in tempo reale.
Al termine troverai l'URL pubblico del tipo:
```
https://dashboard-meteo-XXXXXXXXXX-ey.a.run.app
```

---

## 🔄 Workflow completo

```
CSV mensili dalla stazione
        │
        ▼
┌───────────────────┐
│  Excel + Macro    │  → Dashboard Excel + export_streamlit.json
│  (macro_meteo.bas)│
└───────────────────┘
        │
        ▼ (git push)
┌───────────────────┐
│     GitHub        │  → trigger automatico GitHub Actions
│  (repository)     │
└───────────────────┘
        │
        ▼ (deploy automatico)
┌───────────────────┐
│  Google Cloud Run │  → URL pubblico accessibile ovunque
│  (Streamlit app)  │
└───────────────────┘
```

---

## 💰 Costi Google Cloud

- **Cloud Run**: gratuito fino a 2 milioni di richieste/mese (Free Tier)
- **Cloud Build**: gratuito fino a 120 min/giorno di build
- **Container Registry**: ~$0.02/GB/mese di storage immagini

Per un dashboard ad uso interno, i costi sono tipicamente **$0** o pochi centesimi al mese.

---

## 🔧 Personalizzazioni

### Aggiungere nuove stazioni all'app Streamlit
Modifica `streamlit/app.py` — nella sidebar aggiungi una nuova opzione e il codice per caricare i dati della stazione aggiuntiva.

### Cambiare region Cloud Run
Nel file `.github/workflows/deploy.yml`, modifica la variabile `REGION`:
```yaml
REGION: europe-west1    # Belgio
REGION: europe-west8    # Milano (default)
REGION: us-central1     # Iowa
```

### Dominio personalizzato
In Google Cloud Console → Cloud Run → seleziona il servizio → **Gestisci domini personalizzati**.
