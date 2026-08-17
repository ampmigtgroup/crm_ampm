#!/usr/bin/env python3
from pathlib import Path
from datetime import datetime, timezone
import json, shutil, sys

ROOT = Path.cwd()
CONF = ROOT / "src-tauri" / "tauri.conf.json"

if not CONF.exists():
    print("Execute este script na raiz do projeto.")
    raise SystemExit(2)

conf = json.loads(CONF.read_text(encoding="utf-8"))
version = str(conf.get("version", "")).strip()
if version != "5.3.1":
    print(f"Versão atual inesperada: {version!r}. Esperado: 5.3.1")
    raise SystemExit(3)

tag = f"desktop-v{version}"
release_dir = ROOT / "release" / tag
release_dir.mkdir(parents=True, exist_ok=True)

appimage_dir = ROOT / "src-tauri" / "target" / "release" / "bundle" / "appimage"
appimages = list(appimage_dir.glob(f"*{version}*.AppImage"))
if not appimages:
    print("AppImage V5.3.1 não encontrado.")
    print("Rode primeiro o build assinado.")
    raise SystemExit(4)

appimage = appimages[0]
sig = Path(str(appimage) + ".sig")
if not sig.exists():
    print("Assinatura não encontrada:", sig)
    raise SystemExit(5)

# Nome sem espaços facilita a URL da release.
safe_name = f"CRM_Operacional_AmPm_{version}_amd64.AppImage"
safe_sig_name = safe_name + ".sig"

shutil.copy2(appimage, release_dir / safe_name)
shutil.copy2(sig, release_dir / safe_sig_name)

signature = sig.read_text(encoding="utf-8").strip()

latest = {
    "version": version,
    "notes": (
        "Atualização de validação do sistema de atualização automática "
        "do CRM Operacional AmPm Desktop."
    ),
    "pub_date": (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    ),
    "platforms": {
        "linux-x86_64": {
            "signature": signature,
            "url": (
                "https://github.com/ampmigtgroup/crm_ampm/releases/download/"
                f"{tag}/{safe_name}"
            )
        }
    }
}

(release_dir / "latest.json").write_text(
    json.dumps(latest, indent=2, ensure_ascii=False) + "\n",
    encoding="utf-8"
)

print("✅ Release V5.3.1 preparada")
print("Pasta:", release_dir)
print("Arquivos:")
print(" -", safe_name)
print(" -", safe_sig_name)
print(" - latest.json")
print("")
print("Tag GitHub:", tag)
