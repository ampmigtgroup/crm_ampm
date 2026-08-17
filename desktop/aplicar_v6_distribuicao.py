#!/usr/bin/env python3
from pathlib import Path
import json, re, shutil

ROOT = Path.cwd()
SRC = ROOT / "src-tauri"
CONF = SRC / "tauri.conf.json"
CARGO = SRC / "Cargo.toml"
PACKAGE = ROOT / "package.json"
ABOUT = ROOT / "dist" / "about.html"
SPLASH = ROOT / "dist" / "splash.html"

for p in [CONF, CARGO, PACKAGE]:
    if not p.exists():
        print("Execute este script na raiz do projeto desktop.")
        print("Arquivo ausente:", p)
        raise SystemExit(2)

conf = json.loads(CONF.read_text(encoding="utf-8"))
updater = conf.get("plugins", {}).get("updater", {})
if not updater.get("pubkey") or not updater.get("endpoints"):
    print("Updater ainda não está configurado corretamente.")
    raise SystemExit(3)

backup = ROOT / "backup_antes_v6"
backup.mkdir(exist_ok=True)
for p in [CONF, CARGO, PACKAGE, ABOUT, SPLASH, SRC/"tauri.linux.conf.json", SRC/"tauri.windows.conf.json"]:
    if p.exists():
        shutil.copy2(p, backup / p.name)

package = json.loads(PACKAGE.read_text(encoding="utf-8"))
package["version"] = "6.0.0"
package["description"] = "CRM Operacional AmPm - Desktop Corporativo V6"
PACKAGE.write_text(json.dumps(package, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

cargo = CARGO.read_text(encoding="utf-8")
cargo = re.sub(r'(?m)^version = "[^"]+"', 'version = "6.0.0"', cargo, count=1)
CARGO.write_text(cargo, encoding="utf-8")

conf["version"] = "6.0.0"
conf["productName"] = "CRM Operacional AmPm"
bundle = conf.setdefault("bundle", {})
bundle["active"] = True
bundle["createUpdaterArtifacts"] = True
bundle["publisher"] = "IGT Group"
bundle["category"] = "Business"
bundle["shortDescription"] = "CRM Operacional AmPm"
bundle["longDescription"] = "Aplicativo corporativo para gestão operacional AmPm, Call Center, Pipeline, treinamentos e administração."
bundle["copyright"] = "Copyright © 2026 IGT Group"
conf.setdefault("plugins", {}).setdefault("updater", {}).setdefault("windows", {})["installMode"] = "passive"

CONF.write_text(json.dumps(conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

linux_conf = {
    "bundle": {
        "targets": ["appimage", "deb"],
        "linux": {
            "appimage": {
                "bundleMediaFramework": True
            }
        }
    }
}
(SRC / "tauri.linux.conf.json").write_text(json.dumps(linux_conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

windows_conf = {
    "bundle": {
        "targets": ["nsis"],
        "publisher": "IGT Group",
        "windows": {
            "webviewInstallMode": {
                "type": "downloadBootstrapper"
            },
            "nsis": {
                "installMode": "currentUser",
                "languages": ["PortugueseBR", "English"],
                "displayLanguageSelector": False,
                "installerIcon": "icons/icon.ico"
            }
        }
    }
}
(SRC / "tauri.windows.conf.json").write_text(json.dumps(windows_conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

if ABOUT.exists():
    s = ABOUT.read_text(encoding="utf-8")
    s = re.sub(r'(<b>Versão</b><span>)[^<]+(</span>)', r'\g<1>6.0.0\2', s, count=1)
    ABOUT.write_text(s, encoding="utf-8")

if SPLASH.exists():
    s = SPLASH.read_text(encoding="utf-8")
    s = re.sub(r'Desktop(?: Executivo| Resiliente| Corporativo)? V[0-9.]+', 'Desktop Corporativo V6.0', s)
    SPLASH.write_text(s, encoding="utf-8")

print("✅ V6 aplicada.")
print("✅ Updater e chave pública preservados.")
print("✅ Linux: AppImage + .deb + GStreamer.")
print("✅ Windows: NSIS setup.exe em PortuguêsBR/English.")
print("✅ Windows: instalação currentUser.")
print("✅ Updater Windows: modo passive.")
print("✅ Backup criado em backup_antes_v6/")
