#!/usr/bin/env python3
from pathlib import Path
import json

ROOT = Path.cwd()
SRC = ROOT / "src-tauri"
main = json.loads((SRC/"tauri.conf.json").read_text(encoding="utf-8"))
linux = json.loads((SRC/"tauri.linux.conf.json").read_text(encoding="utf-8"))
windows = json.loads((SRC/"tauri.windows.conf.json").read_text(encoding="utf-8"))

checks = [
    ("versão", main.get("version") == "6.0.0"),
    ("updater pubkey", bool(main.get("plugins", {}).get("updater", {}).get("pubkey"))),
    ("updater endpoint", bool(main.get("plugins", {}).get("updater", {}).get("endpoints"))),
    ("update artifacts", main.get("bundle", {}).get("createUpdaterArtifacts") is True),
    ("Linux AppImage", "appimage" in linux.get("bundle", {}).get("targets", [])),
    ("Linux deb", "deb" in linux.get("bundle", {}).get("targets", [])),
    ("GStreamer AppImage", linux.get("bundle", {}).get("linux", {}).get("appimage", {}).get("bundleMediaFramework") is True),
    ("Windows NSIS", "nsis" in windows.get("bundle", {}).get("targets", [])),
    ("Windows currentUser", windows.get("bundle", {}).get("windows", {}).get("nsis", {}).get("installMode") == "currentUser"),
    ("Windows PortugueseBR", "PortugueseBR" in windows.get("bundle", {}).get("windows", {}).get("nsis", {}).get("languages", [])),
]
failed = False
for name, ok in checks:
    print(("✅" if ok else "❌"), name)
    failed = failed or (not ok)
raise SystemExit(1 if failed else 0)
