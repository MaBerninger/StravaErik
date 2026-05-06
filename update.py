"""
update.py — Strava Dashboard Updater
Exporteert Excel naar JSON en pusht naar GitHub Pages.

Gebruik: python update.py
"""

import subprocess
import sys
import os

# ── Instellingen — pas dit aan ────────────────────────
GITHUB_TOKEN = "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"
GITHUB_USER  = "MaBerninger"
REPO_NAME    = "StravaErik"
MAP          = r"C:\Strava\17"
# ──────────────────────────────────────────────────────

def run(cmd, cwd=None):
    result = subprocess.run(cmd, cwd=cwd, shell=True, capture_output=True, text=True)
    if result.stdout: print(result.stdout.strip())
    if result.stderr: print(result.stderr.strip())
    return result.returncode
 
def main():
    print("=" * 40)
    print("  Strava Dashboard Updater")
    print("=" * 40)
 
    os.chdir(MAP)
 
    # Stap 1: Export
    print("\n[1/3] Excel exporteren naar JSON...")
    if run("python export_strava.py") != 0:
        print("FOUT: export mislukt!")
        sys.exit(1)
 
    # Stap 2: Git init
    print("\n[2/3] Git voorbereiden...")
    remote_url = f"https://{GITHUB_TOKEN}@github.com/{GITHUB_USER}/{REPO_NAME}.git"
 
    if not os.path.exists(os.path.join(MAP, ".git")):
        run("git init")
        run(f"git remote add origin {remote_url}")
        run("git branch -M main")
    else:
        run(f"git remote set-url origin {remote_url}")
 
    run("git add .")
 
 
    # Stap 3: Commit & push
    print("\n[3/3] Pushen naar GitHub...")
    from datetime import datetime
    tijdstip = datetime.now().strftime("%Y-%m-%d %H:%M")
    run(f'git commit -m "Update strava data {tijdstip}"')
    run("git push -u origin main --force")
 
    print("\n" + "=" * 40)
    print("  Klaar! Site bijgewerkt op:")
    print(f"  https://{GITHUB_USER}.github.io/{REPO_NAME}/erik_loopdata.html")
    print("=" * 40)
 
if __name__ == "__main__":
    main()