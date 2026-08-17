#!/usr/bin/env python3
from pathlib import Path
import json, re, shutil, sys

ROOT = Path.cwd()
LIB = ROOT / "src-tauri" / "src" / "lib.rs"
CARGO = ROOT / "src-tauri" / "Cargo.toml"
CONF = ROOT / "src-tauri" / "tauri.conf.json"
PACKAGE = ROOT / "package.json"
ABOUT = ROOT / "dist" / "about.html"
SPLASH = ROOT / "dist" / "splash.html"

required = [LIB, CARGO, CONF, PACKAGE]
missing = [str(p) for p in required if not p.exists()]
if missing:
    print("Execute este script na raiz do projeto desktop.")
    print("Arquivos ausentes:", *missing, sep="\n- ")
    raise SystemExit(2)

conf = json.loads(CONF.read_text(encoding="utf-8"))
updater = conf.get("plugins", {}).get("updater", {})
if not updater.get("pubkey") or not updater.get("endpoints"):
    print("O updater ainda não possui chave pública + endpoint válidos.")
    print("Configure primeiro o updater da V5.1.")
    raise SystemExit(3)

# Backup
backup = ROOT / "backup_v5_3"
backup.mkdir(exist_ok=True)
for p in [LIB, CARGO, CONF, PACKAGE]:
    shutil.copy2(p, backup / p.name)

# O arquivo lib.rs novo deve estar ao lado deste script.
new_lib = Path(__file__).resolve().with_name("lib_v5_3_updater_automatico.rs")
if not new_lib.exists():
    print("lib_v5_3_updater_automatico.rs não encontrado ao lado do script.")
    raise SystemExit(4)

LIB.write_text(new_lib.read_text(encoding="utf-8"), encoding="utf-8")

cargo = CARGO.read_text(encoding="utf-8")
cargo = re.sub(r'(?m)^version = "[^"]+"', 'version = "5.3.0"', cargo, count=1)
if not re.search(r'(?m)^serde_json\s*=', cargo):
    cargo = cargo.replace('[dependencies]\n', '[dependencies]\nserde_json = "1"\n', 1)
if 'tauri-plugin-updater = "2"' not in cargo:
    cargo += '\ntauri-plugin-updater = "2"\n'
CARGO.write_text(cargo, encoding="utf-8")

package = json.loads(PACKAGE.read_text(encoding="utf-8"))
package["version"] = "5.3.0"
PACKAGE.write_text(json.dumps(package, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

# Preserva pubkey/endpoints já existentes.
conf["version"] = "5.3.0"
conf.setdefault("bundle", {})["createUpdaterArtifacts"] = True
CONF.write_text(json.dumps(conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

if ABOUT.exists():
    s = ABOUT.read_text(encoding="utf-8")
    s = re.sub(r'(<b>Versão</b><span>)[^<]+(</span>)', r'\g<1>5.3.0\2', s, count=1)
    ABOUT.write_text(s, encoding="utf-8")

if SPLASH.exists():
    s = SPLASH.read_text(encoding="utf-8")
    s = re.sub(r'Desktop Corporativo V5(?:\.\d+)?', 'Desktop Corporativo V5.3', s)
    SPLASH.write_text(s, encoding="utf-8")

print("✅ V5.3 aplicada.")
print("✅ Chave pública e endpoint foram preservados.")
print("✅ Atualização automática habilitada em builds release.")
print("✅ Menu 'Verificar atualizações' adicionado.")
print("✅ Backup criado em backup_v5_3/")
print("")
print("Agora execute:")
print("  npm run dev")
