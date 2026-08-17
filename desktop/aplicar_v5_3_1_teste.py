#!/usr/bin/env python3
from pathlib import Path
import json, re, shutil, sys

ROOT = Path.cwd()
CARGO = ROOT / "src-tauri" / "Cargo.toml"
CONF = ROOT / "src-tauri" / "tauri.conf.json"
PACKAGE = ROOT / "package.json"
ABOUT = ROOT / "dist" / "about.html"
SPLASH = ROOT / "dist" / "splash.html"

required = [CARGO, CONF, PACKAGE]
missing = [p for p in required if not p.exists()]
if missing:
    print("Execute este script na raiz do projeto desktop.")
    for p in missing:
        print("Ausente:", p)
    raise SystemExit(2)

conf = json.loads(CONF.read_text(encoding="utf-8"))
updater = conf.get("plugins", {}).get("updater", {})

if not updater.get("pubkey") or not updater.get("endpoints"):
    print("Updater não configurado. A chave pública e o endpoint precisam existir.")
    raise SystemExit(3)

# Backup da V5.3.0 antes do bump.
backup = ROOT / "backup_antes_v5_3_1"
backup.mkdir(exist_ok=True)
for p in [CARGO, CONF, PACKAGE, ABOUT, SPLASH]:
    if p.exists():
        shutil.copy2(p, backup / p.name)

# Cargo
cargo = CARGO.read_text(encoding="utf-8")
cargo = re.sub(r'(?m)^version = "[^"]+"', 'version = "5.3.1"', cargo, count=1)
CARGO.write_text(cargo, encoding="utf-8")

# package.json
package = json.loads(PACKAGE.read_text(encoding="utf-8"))
package["version"] = "5.3.1"
PACKAGE.write_text(json.dumps(package, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

# tauri.conf.json - preserva pubkey/endpoints e updater.
conf["version"] = "5.3.1"
conf.setdefault("bundle", {})["createUpdaterArtifacts"] = True
CONF.write_text(json.dumps(conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

# Sobre
if ABOUT.exists():
    s = ABOUT.read_text(encoding="utf-8")
    s = re.sub(
        r'(<b>Versão</b><span>)[^<]+(</span>)',
        r'\g<1>5.3.1\2',
        s,
        count=1
    )
    ABOUT.write_text(s, encoding="utf-8")

# Splash
if SPLASH.exists():
    s = SPLASH.read_text(encoding="utf-8")
    s = re.sub(
        r'Desktop Corporativo V5(?:\.\d+)?',
        'Desktop Corporativo V5.3.1',
        s
    )
    SPLASH.write_text(s, encoding="utf-8")

print("✅ Projeto atualizado para V5.3.1")
print("✅ Chave pública preservada")
print("✅ Endpoint do updater preservado")
print("✅ createUpdaterArtifacts = true")
print("✅ Backup criado em backup_antes_v5_3_1/")
print("")
print("Próximo passo:")
print('  grep \'"version"\' package.json')
print('  grep \'"version"\' src-tauri/tauri.conf.json')
print("  npm run dev")
