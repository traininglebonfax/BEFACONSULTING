"""
Verification avant demonstration client : controle l'environnement puis
deroule la chaine complete par le VRAI serveur (memes routes que le formulaire).

    python -m scripts.verifier_demo          # rapide (~1 min) : environnement, cle API, serveur, PDF sans IA
    python -m scripts.verifier_demo --ia     # repetition complete (~5 min, consomme du credit API) :
                                             # + import de liasse reelle + redaction IA avec recherche web
                                             # + Business Plan final avec redaction

Les documents produits sont ranges dans sorties/verification_demo/ (a ouvrir et relire avant le rendez-vous).
Code de sortie 0 = pret pour la demo ; 1 = au moins un point bloquant.
"""
import importlib
import json
import os
import socket
import subprocess
import sys
import time

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SORTIE = os.path.join(RACINE, "sorties", "verification_demo")
LIASSE_TEST = os.path.join(RACINE, "Model Business Plan", "Bilan 2025 ETTS  .pdf")
TRENDS_TEST = os.path.join(RACINE, "tests", "fixtures", "trends_multiTimeline_huile_de_palme.csv")

resultats = []  # (ok, bloquant, libelle, detail)


def note(ok, libelle, detail="", bloquant=True):
    resultats.append((ok, bloquant, libelle, detail))
    marque = "OK  " if ok else ("ECHEC" if bloquant else "ATTN")
    print(f"[{marque}] {libelle}" + (f" — {detail}" if detail else ""), flush=True)
    return ok


# ----------------------------------------------------------------------
# 1. Environnement
# ----------------------------------------------------------------------
def verifier_environnement():
    print("\n== 1. Environnement ==")
    v = sys.version_info
    note(v >= (3, 10), "Python", f"{v.major}.{v.minor}.{v.micro} ({sys.executable})")
    manquants = []
    for mod in ("fastapi", "uvicorn", "pydantic", "openpyxl", "anthropic", "pypdf", "docx",
                "multipart", "PIL", "matplotlib", "dotenv", "httpx2"):
        try:
            importlib.import_module(mod)
        except ImportError:
            manquants.append(mod)
    note(not manquants, "Dependances Python", "manquantes : " + ", ".join(manquants) if manquants
         else "toutes installees")

    from engine.recalc import find_soffice
    try:
        note(True, "LibreOffice (conversion PDF)", find_soffice())
    except Exception as e:  # noqa: BLE001
        note(False, "LibreOffice (conversion PDF)", f"introuvable : {e}")

    for chemin in ("assets/branding/logo_befa.png", "web/index.html",
                   "data/befa_exemple_coreal.json", "data/befa_exemple_oivoire.json"):
        note(os.path.exists(os.path.join(RACINE, chemin)), f"Fichier {chemin}")


# ----------------------------------------------------------------------
# 2. Cle API Anthropic (cle valable 24 h : a renouveler le jour de la demo)
# ----------------------------------------------------------------------
def verifier_cle_api():
    print("\n== 2. Cle API Anthropic ==")
    from dotenv import load_dotenv
    load_dotenv(os.path.join(RACINE, ".env"), override=True)
    cle = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not note(bool(cle), "ANTHROPIC_API_KEY presente dans .env"):
        return False
    import anthropic
    from befa.ia import modele
    try:
        r = anthropic.Anthropic(timeout=60).messages.create(
            model=modele(), max_tokens=16,
            messages=[{"role": "user", "content": "Réponds uniquement : OK"}])
        return note(True, "Appel test a l'API", f"modele {r.model} repond")
    except anthropic.AuthenticationError:
        return note(False, "Appel test a l'API", "cle invalide ou expiree -> generer une nouvelle cle et "
                    "mettre a jour .env")
    except anthropic.BadRequestError as e:
        msg = str(e)
        if "credit" in msg.lower():
            return note(False, "Appel test a l'API", "credit API epuise -> recharger le compte Anthropic")
        return note(False, "Appel test a l'API", msg[:200])
    except Exception as e:  # noqa: BLE001
        return note(False, "Appel test a l'API", f"{type(e).__name__} : {str(e)[:200]}")


# ----------------------------------------------------------------------
# 3. Serveur reel + parcours du formulaire
# ----------------------------------------------------------------------
def port_libre():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def demarrer_serveur(port):
    journal = open(os.path.join(SORTIE, "serveur.log"), "w", encoding="utf-8")
    proc = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "backend.app:app", "--host", "127.0.0.1", "--port", str(port)],
        cwd=RACINE, stdout=journal, stderr=subprocess.STDOUT)
    import httpx2 as httpx
    for _ in range(60):
        if proc.poll() is not None:
            break
        try:
            if httpx.get(f"http://127.0.0.1:{port}/api/health", timeout=2).status_code == 200:
                return proc
        except httpx.HTTPError:
            time.sleep(0.5)
    proc.kill()
    return None


def telecharger(client, url, nom, entete):
    r = client.get(url)
    ok = r.status_code == 200 and r.content.startswith(entete)
    if ok:
        with open(os.path.join(SORTIE, nom), "wb") as f:
            f.write(r.content)
    return ok, len(r.content)


def generer(client, dossier, nom, redaction=None, trends=True):
    donnees = {"dossier": json.dumps(dossier, ensure_ascii=False)}
    if redaction:
        donnees["redaction"] = json.dumps(redaction, ensure_ascii=False)
    fichiers = []
    if trends:
        fichiers = [("fichiers_trends", ("multiTimeline.csv", open(TRENDS_TEST, "rb").read(), "text/csv"))]
    t = time.time()
    r = client.post("/api/befa/generer", data=donnees, files=fichiers or None)
    duree = time.time() - t
    if r.status_code != 200:
        return note(False, f"Generation {nom}", f"HTTP {r.status_code} : {r.text[:300]}")
    res = r.json()
    ok_pdf, taille_pdf = telecharger(client, res["pdf"], f"{nom}.pdf", b"%PDF")
    ok_docx, _ = telecharger(client, res["docx"], f"{nom}.docx", b"PK")
    note(ok_pdf and ok_docx, f"Generation {nom}",
         f"{res['pages']} pages, {taille_pdf // 1024} Ko, {duree:.0f} s, sommaire "
         f"{'complet' if res['sommaire_complet'] else 'INCOMPLET'}")
    note(res["sommaire_complet"], f"Sommaire {nom}", bloquant=False)
    for a in res.get("alertes", []):
        note(False, f"Alerte {nom}", a, bloquant=False)
    return ok_pdf and ok_docx


def parcours(port, avec_ia):
    import httpx2 as httpx
    client = httpx.Client(base_url=f"http://127.0.0.1:{port}", timeout=900)

    print("\n== 3. Pages et services du formulaire ==")
    r = client.get("/")
    note(r.status_code == 200 and "BEFA" in r.text, "Page d'accueil (formulaire 8 etapes)")
    r = client.get("/assets/branding/logo_befa.png")
    note(r.status_code == 200 and r.content.startswith(b"\x89PNG"), "Logo BEFA")
    coreal = client.get("/api/befa/exemples/coreal").json()
    oivoire = client.get("/api/befa/exemples/oivoire").json()
    note(coreal.get("variante") == "nouveau_projet", "Exemple COREAL (nouveau projet)")
    note(oivoire.get("variante") == "entreprise_existante", "Exemple O'IVOIRE (entreprise existante)")
    r = client.post("/api/befa/projection", json={"projection": coreal["projection"], "nom": "COREAL FOOD"}).json()
    note(r.get("erreurs") == [] and r["synthese"]["cout_investissements"] == 2_700_702_260,
         "Apercu financier recalcule", "investissement COREAL = 2 700 702 260 FCFA")
    with open(TRENDS_TEST, "rb") as f:
        r = client.post("/api/befa/trends", files=[("fichiers", ("multiTimeline.csv", f.read(), "text/csv"))])
    note(r.status_code == 200, "Import Google Trends (CSV)", f"HTTP {r.status_code}")

    print("\n== 4. Business Plan sans IA (secours si l'API tombe pendant la demo) ==")
    generer(client, coreal, "1_COREAL_sans_IA")
    generer(client, oivoire, "2_OIVOIRE_sans_IA", trends=False)

    if not avec_ia:
        return
    print("\n== 5. Repetition complete avec IA (consomme du credit API) ==")
    if os.path.exists(LIASSE_TEST):
        t = time.time()
        with open(LIASSE_TEST, "rb") as f:
            r = client.post("/api/befa/liasse", files={"fichier": ("Bilan 2025 ETTS.pdf", f.read(), "application/pdf")})
        if r.status_code == 200:
            d = r.json()
            with open(os.path.join(SORTIE, "3_liasse_ETTS.json"), "w", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False, indent=1)
            note(d["controles_ok"], "Import liasse ETTS 2025", f"{time.time() - t:.0f} s, controles des totaux "
                 + ("OK" if d["controles_ok"] else "EN ECART : " + "; ".join(d["alertes"])))
        else:
            note(False, "Import liasse ETTS 2025", f"HTTP {r.status_code} : {r.text[:300]}")
    else:
        note(False, "Import liasse", f"fichier de test absent : {LIASSE_TEST}", bloquant=False)

    print("   Redaction IA avec recherche web en cours (3 a 4 minutes)...", flush=True)
    t = time.time()
    r = client.post("/api/befa/redaction", json=coreal)
    if r.status_code != 200:
        note(False, "Redaction IA COREAL", f"HTTP {r.status_code} : {r.text[:300]}")
        return
    red = r.json()
    with open(os.path.join(SORTIE, "4_redaction_COREAL.json"), "w", encoding="utf-8") as f:
        json.dump(red, f, ensure_ascii=False, indent=1)
    texte = json.dumps(red, ensure_ascii=False)
    note(True, "Redaction IA COREAL", f"{time.time() - t:.0f} s, {len(red.get('sources', []))} sources, "
         f"{texte.count('[S')} renvois [Sx], {len(red.get('alertes', []))} alerte(s)")
    generer(client, coreal, "5_COREAL_avec_IA", redaction=red)


def main():
    avec_ia = "--ia" in sys.argv
    os.makedirs(SORTIE, exist_ok=True)
    sys.path.insert(0, RACINE)
    os.chdir(RACINE)
    debut = time.time()

    verifier_environnement()
    cle_ok = verifier_cle_api()
    if avec_ia and not cle_ok:
        print("\nCle API inutilisable : la repetition avec IA est annulee.")
        avec_ia = False

    port = port_libre()
    proc = demarrer_serveur(port)
    if note(proc is not None, "Demarrage du serveur", f"port {port} (journal : sorties/verification_demo/serveur.log)"):
        try:
            parcours(port, avec_ia)
        finally:
            proc.terminate()
            proc.wait(timeout=10)

    bloquants = [r for r in resultats if not r[0] and r[1]]
    attentions = [r for r in resultats if not r[0] and not r[1]]
    print(f"\n== Bilan ({time.time() - debut:.0f} s) ==")
    print(f"{len(resultats) - len(bloquants) - len(attentions)} OK, {len(attentions)} a surveiller, "
          f"{len(bloquants)} bloquant(s)")
    for _, _, lib, det in bloquants:
        print(f"  BLOQUANT : {lib} — {det}")
    print(f"Documents a relire : {SORTIE}")
    print("PRET POUR LA DEMO" if not bloquants else "PAS PRET : corriger les points bloquants ci-dessus")
    sys.exit(1 if bloquants else 0)


if __name__ == "__main__":
    main()
