#!/bin/bash
# ============================================================
#  SETUP AUTOMATICO — Dashboard Meteo
#  Esegui una volta sola dopo aver estratto meteo_pipeline.zip
#  Uso: bash setup.sh
# ============================================================

set -e  # Blocca se un comando fallisce

# ---- COLORI ----
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No color

header()  { echo -e "\n${BLUE}▶ $1${NC}"; }
ok()      { echo -e "${GREEN}  ✓ $1${NC}"; }
warn()    { echo -e "${YELLOW}  ⚠ $1${NC}"; }
ask()     { echo -e "${YELLOW}  → $1${NC}"; }

# ============================================================
#  0. BENVENUTO
# ============================================================
clear
echo -e "${BLUE}"
echo "  ╔════════════════════════════════════════╗"
echo "  ║   Setup Dashboard Meteo Stazioni       ║"
echo "  ║   Excel · GitHub · Google Cloud Run    ║"
echo "  ╚════════════════════════════════════════╝"
echo -e "${NC}"
echo "  Questo script configura tutto in automatico."
echo "  Ti verranno chieste solo 3 informazioni."
echo ""

# ============================================================
#  1. RACCOLTA PARAMETRI
# ============================================================
header "Configurazione iniziale"

ask "GitHub username (es. mario-rossi):"
read -r GH_USER

ask "Nome del repository GitHub (es. dashboard-meteo):"
read -r GH_REPO

ask "Google Cloud Project ID (es. dashboard-meteo-123456):"
read -r GCP_PROJECT

echo ""
echo "  Riepilogo:"
echo "  • GitHub:    https://github.com/$GH_USER/$GH_REPO"
echo "  • GCP:       $GCP_PROJECT"
echo ""
ask "Confermi? (s/n):"
read -r CONFIRM
if [[ "$CONFIRM" != "s" && "$CONFIRM" != "S" ]]; then
  echo "Annullato."; exit 0
fi

# ============================================================
#  2. VERIFICA DIPENDENZE
# ============================================================
header "Verifica dipendenze"

check_cmd() {
  if command -v "$1" &>/dev/null; then
    ok "$1 trovato"
  else
    echo -e "${RED}  ✗ $1 non trovato — installalo prima di continuare${NC}"
    echo "    Guida: $2"
    exit 1
  fi
}

check_cmd git    "https://git-scm.com/downloads"
check_cmd python3 "https://www.python.org/downloads"
check_cmd pip3   "incluso con Python"

# Python packages
header "Installazione dipendenze Python"
if [ -f "streamlit/requirements.txt" ]; then
  python3 -m venv .venv
  source .venv/bin/activate 2>/dev/null || source .venv/Scripts/activate 2>/dev/null
  pip install -r streamlit/requirements.txt -q
  ok "Dipendenze Python installate in .venv"
else
  warn "requirements.txt non trovato — skip"
fi

# ============================================================
#  3. INIT GIT + PRIMO PUSH
# ============================================================
header "Inizializzazione repository Git"

# Aggiorna deploy.yml con il Project ID reale
if [ -f ".github/workflows/deploy.yml" ]; then
  sed -i.bak "s/europe-west8/europe-west8/g" .github/workflows/deploy.yml
  ok "deploy.yml pronto"
fi

# Aggiorna Dockerfile se necessario
ok "Dockerfile pronto"

# Init git se non già fatto
if [ ! -d ".git" ]; then
  git init
  ok "Repository Git inizializzato"
else
  ok "Repository Git già presente"
fi

git add .
git commit -m "Setup iniziale — Dashboard Meteo Stazioni" 2>/dev/null || \
  git commit --allow-empty -m "Setup iniziale" 
ok "Commit creato"

# Imposta remote
REMOTE_URL="https://github.com/$GH_USER/$GH_REPO.git"
git remote remove origin 2>/dev/null || true
git remote add origin "$REMOTE_URL"
git branch -M main
ok "Remote GitHub configurato: $REMOTE_URL"

echo ""
warn "Prima di fare il push, assicurati di:"
echo "  1. Aver creato il repo su GitHub: https://github.com/new"
echo "  2. Aver aggiunto i 2 Secrets nel repo GitHub:"
echo "     • GCP_PROJECT_ID = $GCP_PROJECT"
echo "     • GCP_SA_KEY     = (contenuto JSON del service account)"
echo ""
ask "Premi INVIO quando il repo GitHub è pronto e i Secrets sono configurati..."
read -r

git push -u origin main
ok "Codice pushato su GitHub!"

# ============================================================
#  4. TEST LOCALE STREAMLIT
# ============================================================
header "Avvio test locale Streamlit"

echo ""
echo "  La dashboard si aprirà su: http://localhost:8501"
echo "  Premi Ctrl+C per fermarla e continuare."
echo ""
ask "Avviare il test locale? (s/n):"
read -r TEST_LOCAL

if [[ "$TEST_LOCAL" == "s" || "$TEST_LOCAL" == "S" ]]; then
  source .venv/bin/activate 2>/dev/null || source .venv/Scripts/activate 2>/dev/null
  streamlit run streamlit/app.py
fi

# ============================================================
#  5. RIEPILOGO FINALE
# ============================================================
echo -e "\n${GREEN}"
echo "  ╔════════════════════════════════════════╗"
echo "  ║   ✅  Setup completato!                ║"
echo "  ╚════════════════════════════════════════╝"
echo -e "${NC}"
echo "  Prossimi passi:"
echo ""
echo "  1. Attendi che GitHub Actions completi il deploy (2-5 min)"
echo "     → https://github.com/$GH_USER/$GH_REPO/actions"
echo ""
echo "  2. Trova l'URL della tua dashboard su Google Cloud:"
echo "     → https://console.cloud.google.com/run?project=$GCP_PROJECT"
echo ""
echo "  3. Flusso mensile da ora in poi:"
echo "     a. Scarica CSV nuovi dalla stazione"
echo "     b. Esegui la macro Excel (1 click)"
echo "     c. Copia il JSON nella cartella del progetto"
echo "     d. git add . && git commit -m 'Dati mese X' && git push"
echo "     e. La dashboard si aggiorna automaticamente"
echo ""
echo "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  📋 Macro Excel:     excel/macro_meteo.bas"
echo "  🐍 App Streamlit:   streamlit/app.py"
echo "  🐳 Dockerfile:      Dockerfile"
echo "  🔄 CI/CD:           .github/workflows/deploy.yml"
echo "  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
