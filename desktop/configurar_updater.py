#!/usr/bin/env python3
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parent
CONF = ROOT / "src-tauri" / "tauri.conf.json"
LIB = ROOT / "src-tauri" / "src" / "lib.rs"

if len(sys.argv) != 2:
    print("Uso:")
    print("  python3 configurar_updater.py /caminho/para/chave-publica")
    raise SystemExit(2)

pub_path = Path(sys.argv[1]).expanduser().resolve()
if not pub_path.exists():
    print(f"Arquivo não encontrado: {pub_path}")
    raise SystemExit(2)

pubkey = pub_path.read_text(encoding="utf-8").strip()
if not pubkey:
    print("A chave pública está vazia.")
    raise SystemExit(2)

# Configuração válida exigida pelo plugin updater.
conf = json.loads(CONF.read_text(encoding="utf-8"))
conf.setdefault("bundle", {})["createUpdaterArtifacts"] = True
conf.setdefault("plugins", {})["updater"] = {
    "pubkey": pubkey,
    "endpoints": [
        "https://github.com/ampmigtgroup/crm_ampm/releases/latest/download/latest.json"
    ],
    "windows": {
        "installMode": "passive"
    }
}
CONF.write_text(json.dumps(conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

# Só registra o plugin depois que existe plugins.updater válido.
lib = LIB.read_text(encoding="utf-8")
plugin_line = "        .plugin(tauri_plugin_updater::Builder::new().build())\n"
single_line = "        .plugin(tauri_plugin_single_instance::init(|app, _args, _cwd| {\n"

if plugin_line not in lib:
    if single_line not in lib:
        print("Não encontrei o ponto de ativação do updater em src-tauri/src/lib.rs.")
        raise SystemExit(3)
    lib = lib.replace(single_line, plugin_line + single_line, 1)
    LIB.write_text(lib, encoding="utf-8")

print("✅ Updater ativado corretamente.")
print("✅ Configuração plugins.updater criada com chave pública e endpoint.")
print("✅ createUpdaterArtifacts = true")
print("✅ Plugin updater registrado no Rust.")
print("")
print("A chave PRIVADA não foi lida nem copiada por este script.")
