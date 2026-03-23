"""
Agent Testeur QA — Version avec mémoire persistante
====================================================
• Capture les vrais screenshots de l'écran Windows
• Pilote souris / clavier via pyautogui
• Apprend entre les sessions grâce à memoire_agent.json
"""

import os
import sys

# Vérifie la clé API avant tout import lourd
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
if not API_KEY:
    print("ERREUR : variable d'environnement ANTHROPIC_API_KEY non définie.")
    print("Crée un fichier .env avec : ANTHROPIC_API_KEY=sk-ant-...")
    print("Puis lance : python -m dotenv run python agent_testeur.py")
    print("Ou définis la variable dans ton terminal avant de lancer.")
    sys.exit(1)

import anthropic
import pyautogui
import base64
import io
import json
import time
from datetime import datetime
from PIL import Image
MODEL      = "claude-opus-4-5"
MAX_TOKENS = 4096
MAX_TURNS  = 300

MEMORY_FILE = os.path.join(os.path.dirname(__file__), "memoire_agent.json")

# ── Correction DPI Windows ────────────────────────────────────
# DOIT être appelé avant tout GetSystemMetrics / pyautogui
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()       # fallback

user32         = ctypes.windll.user32
SCREEN_W_PHYS  = user32.GetSystemMetrics(0)   # résolution physique réelle
SCREEN_H_PHYS  = user32.GetSystemMetrics(1)
# Résolution à envoyer à Claude (réduite pour limiter les tokens)
CLAUDE_W, CLAUDE_H = 1280, 720
# Facteur d'échelle pour convertir les coordonnées Claude → physique
SCALE_X = SCREEN_W_PHYS / CLAUDE_W
SCALE_Y = SCREEN_H_PHYS / CLAUDE_H

# Handle de la fenêtre console (pour la minimiser lors des screenshots)
CONSOLE_HWND = ctypes.windll.kernel32.GetConsoleWindow()

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.3

client = anthropic.Anthropic(api_key=API_KEY)


# ══════════════════════════════════════════════════════════════
#  SYSTÈME DE MÉMOIRE PERSISTANTE
# ══════════════════════════════════════════════════════════════

MEMORY_SCHEMA = {
    "version": 1,
    "sessions": [],           # historique des sessions passées
    "site_knowledge": {},     # UI, boutons, navigation découverts
    "agent_knowledge": {},    # comportements propres à chaque agent
    "excel_structure": {},    # structure du fichier Excel
    "learnings": [],          # apprentissages généraux libres
    "errors_and_fixes": [],   # erreurs rencontrées + solutions trouvées
}

def load_memory() -> dict:
    """Charge la mémoire depuis le fichier JSON, ou crée une mémoire vide."""
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                mem = json.load(f)
            print(f"  [Mémoire] Chargée — {len(mem.get('sessions', []))} session(s) précédente(s)")
            return mem
        except (json.JSONDecodeError, KeyError):
            print("  [Mémoire] Fichier corrompu, nouvelle mémoire créée.")
    else:
        print("  [Mémoire] Première session — mémoire vierge.")
    return dict(MEMORY_SCHEMA)


def save_memory(memory: dict):
    """Sauvegarde la mémoire sur disque."""
    with open(MEMORY_FILE, "w", encoding="utf-8") as f:
        json.dump(memory, f, ensure_ascii=False, indent=2)
    print(f"  [Mémoire] Sauvegardée → {MEMORY_FILE}")


def execute_update_memory(params: dict, memory: dict) -> str:
    """
    Exécute l'outil update_memory appelé par Claude.
    Modifie l'objet memory en place et retourne un message de confirmation.
    """
    category = params.get("category", "")
    key      = params.get("key", "")
    value    = params.get("value", "")

    if category == "session_summary":
        # Ajoute une entrée dans l'historique des sessions
        memory["sessions"].append({
            "date":  datetime.now().strftime("%Y-%m-%d %H:%M"),
            "key":   key,
            "value": value,
        })
    elif category == "learnings":
        # Ajoute un apprentissage libre (évite les doublons exacts)
        entry = f"[{key}] {value}"
        if entry not in memory.setdefault("learnings", []):
            memory["learnings"].append(entry)
    elif category == "errors_and_fixes":
        memory.setdefault("errors_and_fixes", []).append({
            "error": key,
            "fix":   value,
            "date":  datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
    elif category in ("site_knowledge", "agent_knowledge", "excel_structure"):
        memory.setdefault(category, {})[key] = value
    else:
        return f"Catégorie inconnue : {category}"

    save_memory(memory)
    return f"Mémoire mise à jour — [{category}] {key}"


def build_memory_section(memory: dict) -> str:
    """
    Construit le bloc texte à injecter dans le prompt système
    pour que Claude connaisse ses apprentissages précédents.
    """
    if not any([
        memory.get("site_knowledge"),
        memory.get("agent_knowledge"),
        memory.get("excel_structure"),
        memory.get("learnings"),
        memory.get("errors_and_fixes"),
        memory.get("sessions"),
    ]):
        return ""   # première session, rien à injecter

    lines = [
        "",
        "╔══════════════════════════════════════════════════════════════╗",
        "║         MÉMOIRE DES SESSIONS PRÉCÉDENTES                     ║",
        "║  Utilise ces apprentissages pour être plus efficace.         ║",
        "╚══════════════════════════════════════════════════════════════╝",
    ]

    # Historique des sessions
    sessions = memory.get("sessions", [])
    if sessions:
        lines.append("\n── Historique des sessions ─────────────────────────────────")
        for s in sessions[-5:]:   # 5 dernières seulement
            lines.append(f"  • {s['date']} | {s['key']} : {s['value']}")

    # Connaissance du site
    site = memory.get("site_knowledge", {})
    if site:
        lines.append("\n── Connaissance du site web ────────────────────────────────")
        for k, v in site.items():
            lines.append(f"  • {k} : {v}")

    # Connaissance des agents
    agents = memory.get("agent_knowledge", {})
    if agents:
        lines.append("\n── Comportements des agents ────────────────────────────────")
        for k, v in agents.items():
            lines.append(f"  • {k} : {v}")

    # Structure Excel
    excel = memory.get("excel_structure", {})
    if excel:
        lines.append("\n── Structure du fichier Excel ──────────────────────────────")
        for k, v in excel.items():
            lines.append(f"  • {k} : {v}")

    # Apprentissages généraux
    learnings = memory.get("learnings", [])
    if learnings:
        lines.append("\n── Apprentissages généraux ─────────────────────────────────")
        for l in learnings[-20:]:   # 20 derniers
            lines.append(f"  • {l}")

    # Erreurs connues et solutions
    fixes = memory.get("errors_and_fixes", [])
    if fixes:
        lines.append("\n── Erreurs connues et solutions ────────────────────────────")
        for f in fixes[-10:]:
            lines.append(f"  • Erreur : {f['error']}")
            lines.append(f"    Fix    : {f['fix']}")

    lines.append("")
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════
#  DÉFINITION DES OUTILS
# ══════════════════════════════════════════════════════════════

COMPUTER_TOOL = {
    "type": "computer_20250124",
    "name": "computer",
    "display_width_px":  CLAUDE_W,   # Claude raisonne dans cet espace
    "display_height_px": CLAUDE_H,
    "display_number": 1,
}

UPDATE_MEMORY_TOOL = {
    "name": "update_memory",
    "description": (
        "Sauvegarde un apprentissage dans la mémoire persistante. "
        "Appelle cet outil dès que tu découvres quelque chose d'utile "
        "pour les sessions futures : emplacement d'un bouton, temps de "
        "réponse d'un agent, structure de l'Excel, erreur et sa solution, etc. "
        "La mémoire persiste entre les sessions et t'évite de redécouvrir "
        "les mêmes choses à chaque fois."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "category": {
                "type": "string",
                "enum": [
                    "site_knowledge",    # UI, navigation, boutons du site
                    "agent_knowledge",   # comportements propres à un agent
                    "excel_structure",   # colonnes, format du fichier Excel
                    "learnings",         # apprentissage général libre
                    "errors_and_fixes",  # erreur rencontrée + solution
                    "session_summary",   # bilan de la session en cours
                ],
                "description": "Catégorie de l'information à mémoriser.",
            },
            "key": {
                "type": "string",
                "description": (
                    "Identifiant court de l'information. "
                    "Ex: 'bouton_reset', 'Agent Juridique', 'colonne_OK', "
                    "'erreur_timeout', 'bilan_session'."
                ),
            },
            "value": {
                "type": "string",
                "description": (
                    "Description précise de l'information. "
                    "Ex: 'Le bouton Reset se trouve en haut à droite de "
                    "l'interface agent, icône corbeille, coordonnées ~(950, 120).' "
                    "Sois précis pour que tu puisses t'en souvenir facilement."
                ),
            },
        },
        "required": ["category", "key", "value"],
    },
}


# ══════════════════════════════════════════════════════════════
#  PROMPT SYSTÈME (injecté dynamiquement avec la mémoire)
# ══════════════════════════════════════════════════════════════

BASE_SYSTEM_PROMPT = """
Agent QA qui contrôle l'écran réel. Chaque screenshot = état actuel du bureau.

{memory_section}

CONTEXTE : Excel (tests) et Edge (https://elio.onepoint.cloud/consulting-agents) sont déjà ouverts.
Navigation : barre des tâches en bas ou Alt+Tab. Rien d'autre.
Fenêtre noire (terminal) = ignore totalement. Ne jamais l'interagir.
Règle absolue : si une action échoue 2 fois → changer d'approche immédiatement.

--- PHASE 1 : LIRE L'EXCEL ---
Clic Excel dans barre des tâches. Lire tous les onglets (1 onglet = 1 agent).
Pour chaque onglet : mémoriser les tests (description, résultat attendu, colonnes statut/commentaire).
Mémoriser la structure des colonnes une fois pour toutes.

--- PHASE 2 : CARTOGRAPHIER CHAQUE AGENT (avant son premier test) ---
Si la fiche de l'agent est déjà en mémoire → passer directement aux tests.
Sinon, ouvrir l'agent et répondre à ces questions en observant l'écran :

INPUTS : zone de texte (où ?), bouton envoi (libellé, position) ou Entrée seul,
  upload fichier (oui/non), autres paramètres/options visibles, étapes multiples ?

RÉPONSE : où s'affiche-t-elle, quel format (chat/document/tableau/liste),
  streaming progressif ou bloc unique, signal de fin (spinner, bouton réactivé,
  texte stable 3s, indicateur "..."), temps moyen estimé.

RESET entre tests : bouton dédié (libellé, position), ou F5, ou reouvrir l'agent.

PARTICULARITÉS : multi-tours, génère un fichier, demande infos au démarrage,
  boutons sur la réponse (copier/télécharger/régénérer), onglets internes.

Mémoriser la fiche : update_memory(category="agent_knowledge", key="[Nom]_fiche",
  value="inputs|réponse|reset|fin_détection|particularités")
Revenir page d'accueil.

--- PHASE 3 : TESTER ---
Pour chaque agent dans l'ordre des onglets Excel :

  1. Ouvrir l'agent. Utiliser la fiche mémoire.
  2. Pour chaque test :
     a. Reset (méthode de la fiche).
     b. Lire le test dans Excel (Alt+Tab, onglet agent, ligne test, Alt+Tab retour).
     c. Interagir selon la fiche : bon champ, bonne méthode d'envoi,
        uploader fichier si nécessaire, suivre le workflow si multi-étapes.
     d. Attendre fin de réponse avec le signal identifié. Max 90s.
        Screenshots espacés de 5-10s si réponse lente. Juger APRÈS la fin.
     e. Évaluer par rapport au critère Excel (pas mot pour mot, mais intention) :
        OK = critère satisfait, réponse cohérente, pas d'erreur technique.
        KO = critère non satisfait, erreur/crash/timeout/hors-sujet/refus injustifié.
        Si l'agent pose une question → répondre brièvement, noter en commentaire.
     f. Alt+Tab Excel → noter statut OK/KO + commentaire factuel (2 phrases).
        Ctrl+S. Alt+Tab Edge.
  3. Fin d'agent : Ctrl+S Excel, mémoriser bilan, retour page d'accueil.
     update_memory(category="agent_knowledge", key="[Nom]_bilan",
       value="X/Y OK, temps ~Xs, bugs:[...], conseils:[...]")

--- PHASE 4 : RAPPORT FINAL ---
Ctrl+S Excel. Mémoriser bilan session. Produire rapport : totaux OK/KO,
par agent taux + bugs, patterns récurrents, recommandations.

Mémoriser au fur et à mesure (pas à la fin) : interfaces, positions icônes,
colonnes Excel, timings, erreurs+solutions.

Commence. Screenshot.
"""


def build_system_prompt(memory: dict) -> str:
    """Injecte le bloc mémoire dans le prompt système."""
    memory_section = build_memory_section(memory)
    return BASE_SYSTEM_PROMPT.format(memory_section=memory_section)


# ══════════════════════════════════════════════════════════════
#  EXÉCUTION DES ACTIONS SUR LE VRAI BUREAU
# ══════════════════════════════════════════════════════════════

def take_screenshot() -> str:
    """
    Capture l'écran réel, minimise d'abord le terminal pour ne pas le voir
    dans le screenshot, puis retourne une image base64 PNG redimensionnée
    à CLAUDE_W×CLAUDE_H pour que les coordonnées Claude soient cohérentes.
    """
    SW_MINIMIZE = 6
    SW_RESTORE  = 9

    # Minimise le terminal — il reste minimisé pendant toute la session
    # (les print() continuent de fonctionner même fenêtre réduite)
    if CONSOLE_HWND:
        ctypes.windll.user32.ShowWindow(CONSOLE_HWND, SW_MINIMIZE)
        time.sleep(0.4)   # laisse Windows redessiner

    screenshot = pyautogui.screenshot()
    # NE PAS restaurer : le terminal reste dans la barre des tâches

    # Redimensionne exactement à CLAUDE_W×CLAUDE_H
    screenshot = screenshot.resize((CLAUDE_W, CLAUDE_H), Image.LANCZOS)

    buf = io.BytesIO()
    screenshot.save(buf, format="PNG", optimize=True)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8")


def scale(cx: int, cy: int) -> tuple[int, int]:
    """Convertit les coordonnées Claude (1280×720) en coordonnées physiques."""
    return int(cx * SCALE_X), int(cy * SCALE_Y)


def execute_computer_action(action: dict) -> dict:
    act = action.get("action", "")

    if act == "screenshot":
        return {"type": "tool_result_image", "data": take_screenshot(), "media_type": "image/png"}

    elif act == "mouse_move":
        cx, cy = action["coordinate"]
        x, y = scale(cx, cy)
        pyautogui.moveTo(x, y, duration=0.2)
        return {"type": "text", "text": f"Souris déplacée vers ({cx},{cy}) → physique ({x},{y})"}

    elif act == "left_click":
        cx, cy = action["coordinate"]
        x, y = scale(cx, cy)
        pyautogui.click(x, y)
        return {"type": "text", "text": f"Clic gauche ({cx},{cy}) → physique ({x},{y})"}

    elif act == "left_click_drag":
        sx, sy = scale(*action["start_coordinate"])
        ex, ey = scale(*action["coordinate"])
        pyautogui.mouseDown(sx, sy)
        time.sleep(0.1)
        pyautogui.moveTo(ex, ey, duration=0.4)
        pyautogui.mouseUp()
        return {"type": "text", "text": f"Glisser ({sx},{sy}) → ({ex},{ey})"}

    elif act == "right_click":
        cx, cy = action["coordinate"]
        x, y = scale(cx, cy)
        pyautogui.rightClick(x, y)
        return {"type": "text", "text": f"Clic droit ({cx},{cy}) → physique ({x},{y})"}

    elif act == "middle_click":
        cx, cy = action["coordinate"]
        x, y = scale(cx, cy)
        pyautogui.middleClick(x, y)
        return {"type": "text", "text": f"Clic milieu ({cx},{cy}) → physique ({x},{y})"}

    elif act == "double_click":
        cx, cy = action["coordinate"]
        x, y = scale(cx, cy)
        pyautogui.doubleClick(x, y)
        return {"type": "text", "text": f"Double-clic ({cx},{cy}) → physique ({x},{y})"}

    elif act == "type":
        text = action.get("text", "")
        import pyperclip
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        return {"type": "text", "text": f"Texte saisi : {repr(text[:80])}"}

    elif act == "key":
        key_str = action.get("text", "")
        parts = key_str.lower().replace("super", "win").split("+")
        parts = [p.strip() for p in parts]
        if len(parts) == 1:
            pyautogui.press(parts[0])
        else:
            pyautogui.hotkey(*parts)
        return {"type": "text", "text": f"Touche(s) : {key_str}"}

    elif act == "scroll":
        cx, cy    = action["coordinate"]
        x, y      = scale(cx, cy)
        direction  = action.get("direction", "down")
        amount     = int(action.get("amount", 3))
        pyautogui.moveTo(x, y)
        clicks = amount if direction == "up" else -amount
        pyautogui.scroll(clicks)
        return {"type": "text", "text": f"Scroll {direction} x{amount} en ({cx},{cy})"}

    elif act == "wait":
        ms = int(action.get("duration_ms", 2000))
        time.sleep(ms / 1000)
        return {"type": "text", "text": f"Attente {ms}ms"}

    elif act == "cursor_position":
        px, py = pyautogui.position()
        # Retourne en coordonnées Claude pour cohérence
        cx, cy = int(px / SCALE_X), int(py / SCALE_Y)
        return {"type": "text", "text": f"Position curseur : ({cx},{cy}) [physique: ({px},{py})]"}

    else:
        return {"type": "text", "text": f"Action inconnue ignorée : {act}"}


# ══════════════════════════════════════════════════════════════
#  BOUCLE AGENTIQUE PRINCIPALE
# ══════════════════════════════════════════════════════════════

def run_agent():
    # ── Chargement de la mémoire ─────────────────────────────
    memory = load_memory()

    print("═" * 62)
    print("  AGENT TESTEUR QA — avec mémoire persistante")
    print("═" * 62)
    print(f"  Résolution physique : {SCREEN_W_PHYS}×{SCREEN_H_PHYS}")
    print(f"  Espace Claude       : {CLAUDE_W}×{CLAUDE_H}  (scale ×{SCALE_X:.2f})")
    print(f"  Début               : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"  Sessions précédentes : {len(memory.get('sessions', []))}")
    print(f"  Apprentissages en mémoire : {len(memory.get('learnings', []))}")
    print("═" * 62)
    print()
    print("⚠️  Coin supérieur gauche de l'écran = arrêt d'urgence")
    print()
    input("  Appuie sur ENTRÉE pour lancer l'agent...")
    print()

    system_prompt = build_system_prompt(memory)

    print("[Init] Capture du screenshot initial...")
    img_b64 = take_screenshot()

    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": (
                        "Lance la mission de test QA. Consulte d'abord ta mémoire "
                        "des sessions précédentes (injectée dans ton système de prompt) "
                        "pour t'orienter rapidement. Voici l'état actuel de l'écran."
                    ),
                },
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": img_b64,
                    },
                },
            ],
        }
    ]

    turn = 0

    while turn < MAX_TURNS:
        turn += 1
        print(f"\n[Tour {turn}/{MAX_TURNS}] Appel API...")

        try:
            response = client.beta.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=system_prompt,
                tools=[COMPUTER_TOOL, UPDATE_MEMORY_TOOL],
                messages=messages,
                betas=["computer-use-2025-01-24"],
            )
        except anthropic.APIError as e:
            print(f"  ❌ Erreur API : {e}")
            break

        stop_reason = response.stop_reason
        print(f"  Stop reason : {stop_reason}")

        for block in response.content:
            if block.type == "text":
                print("\n── Agent ───────────────────────────────────────────")
                print(block.text)
                print("────────────────────────────────────────────────────")
            elif block.type == "tool_use":
                inp = block.input
                if block.name == "update_memory":
                    print(f"  [Mémoire] [{inp.get('category')}] {inp.get('key')} → {inp.get('value', '')[:80]}")
                else:
                    act   = inp.get("action", "?")
                    coord = inp.get("coordinate", "")
                    txt   = inp.get("text", "")
                    print(f"  [Bureau] {act}"
                          + (f" @ {coord}" if coord else "")
                          + (f' → "{txt[:60]}"' if txt else ""))

        if stop_reason == "end_turn":
            print("\n✅ Mission terminée.")
            break

        if stop_reason == "tool_use":
            messages.append({"role": "assistant", "content": response.content})
            tool_results = []

            for block in response.content:
                if block.type != "tool_use":
                    continue

                # ── Outil mémoire ────────────────────────────────
                if block.name == "update_memory":
                    result_text = execute_update_memory(block.input, memory)
                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": result_text,
                    })

                # ── Outil bureau ─────────────────────────────────
                else:
                    result = execute_computer_action(block.input)
                    if result["type"] == "tool_result_image":
                        content = [{
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": result["media_type"],
                                "data": result["data"],
                            },
                        }]
                    else:
                        content = result["text"]

                    tool_results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": content,
                    })

            messages.append({"role": "user", "content": tool_results})

        else:
            print(f"  ⚠️  Stop reason inattendu : {stop_reason}. Arrêt.")
            break

    if turn >= MAX_TURNS:
        print(f"\n⚠️  Limite de {MAX_TURNS} tours atteinte.")
        # Sauvegarde d'urgence de la mémoire
        save_memory(memory)

    print(f"\n  Fin : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("═" * 62)
    print("  SESSION TERMINÉE")
    print("═" * 62)


# ──────────────────────────────────────────────────────────────
#  POINT D'ENTRÉE
# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        run_agent()
    except KeyboardInterrupt:
        print("\n\n⛔ Arrêt manuel (Ctrl+C).")
        sys.exit(0)
    except pyautogui.FailSafeException:
        print("\n\n⛔ ARRÊT D'URGENCE — souris en coin supérieur gauche.")
        sys.exit(0)
