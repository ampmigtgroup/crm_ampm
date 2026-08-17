#!/usr/bin/env python3
from pathlib import Path
import shutil

ROOT = Path.cwd()
LIB = ROOT / "src-tauri" / "src" / "lib.rs"

if not LIB.exists():
    print("Execute este script na raiz do projeto desktop.")
    print("Esperado:", LIB)
    raise SystemExit(2)

code = LIB.read_text(encoding="utf-8")

backup_dir = ROOT / "backup_antes_modo_fechado"
backup_dir.mkdir(exist_ok=True)
shutil.copy2(LIB, backup_dir / "lib.rs")

closed_ui_script = r"""
const CRM_CLOSED_UI_SCRIPT: &str = r#"
(() => {
  const CRM_ORIGIN = 'https://crmampm-operacional.streamlit.app';
  if (window.location.origin !== CRM_ORIGIN) return;

  const STYLE_ID = 'igt-desktop-closed-ui-style';

  function installStyle() {
    if (document.getElementById(STYLE_ID)) return;

    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #MainMenu,
      footer,
      [data-testid="stToolbar"],
      [data-testid="stDecoration"],
      [data-testid="stStatusWidget"],
      [data-testid="stAppDeployButton"],
      [data-testid="stDeployButton"],
      [data-testid="stAppViewerBadge"],
      [data-testid="stAppCreatorAvatar"],
      [data-testid="stHeaderActionElements"],
      [data-testid="stToolbarActions"] {
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
        opacity: 0 !important;
      }
    `;
    (document.head || document.documentElement).appendChild(style);
  }

  function textOf(el) {
    return [
      el.getAttribute && el.getAttribute('href'),
      el.getAttribute && el.getAttribute('title'),
      el.getAttribute && el.getAttribute('aria-label'),
      el.textContent
    ].filter(Boolean).join(' ').toLowerCase();
  }

  function isTechnical(el) {
    const text = textOf(el);
    return (
      text.includes('github.com') ||
      text.includes('streamlit.io') ||
      text.includes('share.streamlit.io') ||
      text.includes('manage app') ||
      text.includes('deploy') ||
      text.includes('streamlit') ||
      text.includes('github')
    );
  }

  function isBottomRightUtility(el) {
    try {
      const r = el.getBoundingClientRect();
      if (!r || !r.width || !r.height) return false;
      const nearRight = (window.innerWidth - r.right) <= 150;
      const nearBottom = (window.innerHeight - r.bottom) <= 150;
      const compact = r.width <= 300 && r.height <= 190;
      const pos = window.getComputedStyle(el).position;
      const floating = ['fixed', 'sticky', 'absolute'].includes(pos);
      return nearRight && nearBottom && compact && floating;
    } catch (_) {
      return false;
    }
  }

  function hide(el) {
    if (!el || !el.style) return;
    el.style.setProperty('display', 'none', 'important');
    el.style.setProperty('visibility', 'hidden', 'important');
    el.style.setProperty('pointer-events', 'none', 'important');
    el.style.setProperty('opacity', '0', 'important');
  }

  function clean() {
    installStyle();

    const selectors = [
      '#MainMenu',
      'footer',
      '[data-testid="stToolbar"]',
      '[data-testid="stDecoration"]',
      '[data-testid="stStatusWidget"]',
      '[data-testid="stAppDeployButton"]',
      '[data-testid="stDeployButton"]',
      '[data-testid="stAppViewerBadge"]',
      '[data-testid="stAppCreatorAvatar"]',
      '[data-testid="stHeaderActionElements"]',
      '[data-testid="stToolbarActions"]'
    ];

    selectors.forEach((s) => {
      document.querySelectorAll(s).forEach(hide);
    });

    document.querySelectorAll('a, button').forEach((el) => {
      if (!isTechnical(el)) return;

      let candidate = el;
      for (let i = 0; i < 6 && candidate; i += 1) {
        if (isBottomRightUtility(candidate)) {
          hide(candidate);
          return;
        }
        candidate = candidate.parentElement;
      }
    });
  }

  function start() {
    clean();

    const observer = new MutationObserver(clean);
    observer.observe(document.documentElement, {
      childList: true,
      subtree: true
    });

    setInterval(clean, 2500);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();
"#;
"""

if "CRM_CLOSED_UI_SCRIPT" not in code:
    resilience_pos = code.find("const CRM_RESILIENCE_SCRIPT")
    if resilience_pos != -1:
        end = code.find('"#;', resilience_pos)
        if end == -1:
            print("Não consegui localizar o final do CRM_RESILIENCE_SCRIPT.")
            raise SystemExit(3)
        end += 3
        code = code[:end] + "\n\n" + closed_ui_script + code[end:]
    else:
        crm_pos = code.find("const CRM_URL:")
        if crm_pos == -1:
            print("Não encontrei CRM_URL no lib.rs.")
            raise SystemExit(4)
        end_line = code.find("\n", crm_pos)
        code = code[:end_line + 1] + "\n" + closed_ui_script + code[end_line + 1:]

init_line = ".initialization_script(CRM_CLOSED_UI_SCRIPT)"
if init_line not in code:
    preferred = ".initialization_script(CRM_RESILIENCE_SCRIPT)"
    if preferred in code:
        code = code.replace(
            preferred,
            preferred + "\n    " + init_line,
            1
        )
    else:
        title = '.title("CRM Operacional AmPm")'
        if title not in code:
            print("Não encontrei o builder da janela principal.")
            raise SystemExit(5)
        code = code.replace(title, title + "\n    " + init_line, 1)

LIB.write_text(code, encoding="utf-8")

print("✅ Modo fechado aplicado.")
print("✅ Streamlit/GitHub/Deploy/Manage app serão ocultados no desktop.")
print("✅ Updater e configuração atual foram preservados.")
print("✅ Backup: backup_antes_modo_fechado/lib.rs")
print("")
print("Agora execute:")
print("  npm run dev")
