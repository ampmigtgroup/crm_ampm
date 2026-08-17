#!/usr/bin/env python3
from pathlib import Path
from datetime import datetime, timezone
import json, shutil, sys, re

ROOT = Path.cwd()
CONF = ROOT / "src-tauri" / "tauri.conf.json"

if not CONF.exists():
    print("Execute na raiz do projeto.")
    raise SystemExit(2)

conf = json.loads(CONF.read_text(encoding="utf-8"))
version = conf.get("version")
if not version:
    print("Versão não encontrada no tauri.conf.json")
    raise SystemExit(2)

tag = f"desktop-v{version}"
release_dir = ROOT / "release" / tag
release_dir.mkdir(parents=True, exist_ok=True)

appimage_dir = ROOT / "src-tauri" / "target" / "release" / "bundle" / "appimage"
matches = list(appimage_dir.glob("*.AppImage"))
if not matches:
    print("AppImage não encontrado. Rode o build assinado primeiro.")
    raise SystemExit(3)

# Escolhe AppImage da versão atual.
appimage = next((p for p in matches if version in p.name), matches[0])
sig = Path(str(appimage) + ".sig")
if not sig.exists():
    print(f"Assinatura não encontrada: {sig}")
    raise SystemExit(4)

safe_name = f"CRM_Operacional_AmPm_{version}_amd64.AppImage"
safe_sig = safe_name + ".sig"

shutil.copy2(appimage, release_dir / safe_name)
shutil.copy2(sig, release_dir / safe_sig)

signature = sig.read_text(encoding="utf-8").strip()

latest = {
    "version": version,
    "notes": f"CRM Operacional AmPm Desktop v{version}",
    "pub_date": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
    "platforms": {
        "linux-x86_64": {
            "signature": signature,
            "url": f"https://github.com/ampmigtgroup/crm_ampm/releases/download/{tag}/{safe_name}"
        }
    }
}

(release_dir / "latest.json").write_text(
    json.dumps(latest, indent=2, ensure_ascii=False) + "\n",
    encoding="utf-8"
)

print("✅ Pasta de release criada:", release_dir)
print("✅", safe_name)
print("✅", safe_sig)
print("✅ latest.json")
print("")
print("Tag prevista:", tag)
print("Faça upload dos 3 arquivos da pasta release na GitHub Release.")
