use tauri::{
    AppHandle, Manager, WebviewUrl,
    menu::{Menu, MenuBuilder, MenuItem, SubmenuBuilder},
    tray::TrayIconBuilder,
    webview::{PageLoadEvent, WebviewWindowBuilder},
};

use tauri_plugin_updater::UpdaterExt;

const CRM_URL: &str = "https://crmampm-operacional.streamlit.app/";


const CRM_RESILIENCE_SCRIPT: &str = r#"
(() => {
  const CRM_ORIGIN = 'https://crmampm-operacional.streamlit.app';

  if (window.location.origin !== CRM_ORIGIN) return;

  const STYLE_ID = 'igt-desktop-network-style';
  const OVERLAY_ID = 'igt-desktop-network-overlay';

  function ensureUI() {
    if (!document.getElementById(STYLE_ID)) {
      const style = document.createElement('style');
      style.id = STYLE_ID;
      style.textContent = `
        #${OVERLAY_ID}{
          position:fixed; inset:0; z-index:2147483647;
          display:none; align-items:center; justify-content:center;
          padding:24px;
          background:
            radial-gradient(circle at 90% 10%, rgba(255,194,14,.18), transparent 26%),
            rgba(246,248,251,.985);
          font-family:Inter,system-ui,-apple-system,"Segoe UI",sans-serif;
        }
        #${OVERLAY_ID}.show{display:flex}
        #${OVERLAY_ID} .igt-card{
          width:min(520px,100%);
          background:white; border:1px solid #e3e8ef; border-radius:24px;
          padding:30px; text-align:center;
          box-shadow:0 20px 60px rgba(32,45,58,.14);
        }
        #${OVERLAY_ID} .igt-icon{
          width:64px;height:64px;border-radius:19px;margin:0 auto 18px;
          display:grid;place-items:center;
          background:linear-gradient(145deg,#F36F21,#FF8A20);
          color:white;font-size:28px;font-weight:900;
          box-shadow:0 14px 30px rgba(243,111,33,.25);
        }
        #${OVERLAY_ID} h2{margin:0 0 8px;color:#222d39;font-size:25px}
        #${OVERLAY_ID} p{margin:0 auto;color:#718091;line-height:1.55;max-width:410px}
        #${OVERLAY_ID} button{
          margin-top:21px;border:0;border-radius:11px;padding:12px 18px;
          background:#245A9B;color:white;font-weight:800;cursor:pointer;
        }
        #${OVERLAY_ID} .igt-status{
          margin-top:15px;color:#98a2ad;font-size:11px;font-weight:750
        }
      `;
      (document.head || document.documentElement).appendChild(style);
    }

    if (!document.getElementById(OVERLAY_ID) && document.body) {
      const overlay = document.createElement('div');
      overlay.id = OVERLAY_ID;
      overlay.innerHTML = `
        <div class="igt-card">
          <div class="igt-icon">!</div>
          <h2>Conexão interrompida</h2>
          <p>O CRM continua aberto, mas não consegue alcançar o servidor neste momento. Verifique a internet e tente novamente.</p>
          <button type="button" id="igt-desktop-retry">Tentar reconectar</button>
          <div class="igt-status" id="igt-desktop-net-status">Aguardando conexão…</div>
        </div>`;
      document.body.appendChild(overlay);

      const btn = document.getElementById('igt-desktop-retry');
      if (btn) btn.addEventListener('click', () => checkConnection(true));
    }
  }

  function setOffline(message) {
    ensureUI();
    const overlay = document.getElementById(OVERLAY_ID);
    const status = document.getElementById('igt-desktop-net-status');
    if (status && message) status.textContent = message;
    if (overlay) overlay.classList.add('show');
  }

  function setOnline() {
    const overlay = document.getElementById(OVERLAY_ID);
    if (overlay) overlay.classList.remove('show');
  }

  async function checkConnection(reloadOnSuccess = false) {
    if (!navigator.onLine) {
      setOffline('Sem acesso à rede.');
      return false;
    }

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);

    try {
      await fetch(CRM_ORIGIN, {
        method: 'HEAD',
        cache: 'no-store',
        signal: controller.signal
      });
      clearTimeout(timer);
      setOnline();
      if (reloadOnSuccess) window.location.reload();
      return true;
    } catch (_) {
      clearTimeout(timer);
      setOffline('Servidor indisponível ou conexão instável.');
      return false;
    }
  }

  window.addEventListener('offline', () => setOffline('Sem acesso à rede.'));
  window.addEventListener('online', () => checkConnection(false));

  const boot = () => {
    ensureUI();
    checkConnection(false);
    setInterval(() => checkConnection(false), 15000);
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot, { once: true });
  } else {
    boot();
  }
})();
"#;


fn show_main(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.unminimize();
        let _ = window.show();
        let _ = window.set_focus();
    }
}

fn reload_crm(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.reload();
        let _ = window.show();
        let _ = window.set_focus();
    }
}

fn toggle_fullscreen(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        if let Ok(current) = window.is_fullscreen() {
            let _ = window.set_fullscreen(!current);
        }
    }
}

fn show_about(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("about") {
        let _ = window.show();
        let _ = window.set_focus();
        return;
    }

    let _ = WebviewWindowBuilder::new(
        app,
        "about",
        WebviewUrl::App("about.html".into()),
    )
    .title("Sobre • CRM Operacional AmPm")
    .inner_size(540.0, 520.0)
    .min_inner_size(500.0, 470.0)
    .center()
    .resizable(false)
    .maximizable(false)
    .minimizable(true)
    .build();
}


fn set_main_title(app: &AppHandle, title: &str) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.set_title(title);
    }
}

fn check_for_updates(app: AppHandle, manual: bool) {
    tauri::async_runtime::spawn(async move {
        // Em desenvolvimento evitamos checagem automática.
        // A checagem manual continua disponível pelo menu.
        if cfg!(debug_assertions) && !manual {
            return;
        }

        if manual {
            set_main_title(&app, "CRM Operacional AmPm • Verificando atualização...");
        }

        let updater = match app.updater() {
            Ok(updater) => updater,
            Err(err) => {
                eprintln!("Falha ao inicializar updater: {err}");
                if manual {
                    set_main_title(
                        &app,
                        "CRM Operacional AmPm • Falha ao verificar atualização",
                    );
                }
                return;
            }
        };

        match updater.check().await {
            Ok(Some(update)) => {
                let target_version = update.version.clone();

                set_main_title(
                    &app,
                    &format!(
                        "CRM Operacional AmPm • Atualizando para v{}...",
                        target_version
                    ),
                );

                let result = update
                    .download_and_install(
                        |_chunk_length, _content_length| {},
                        || {},
                    )
                    .await;

                match result {
                    Ok(()) => {
                        set_main_title(
                            &app,
                            &format!(
                                "CRM Operacional AmPm • Atualização v{} instalada",
                                target_version
                            ),
                        );
                        app.request_restart();
                    }
                    Err(err) => {
                        eprintln!("Falha ao baixar/instalar atualização: {err}");
                        set_main_title(
                            &app,
                            "CRM Operacional AmPm • Falha ao instalar atualização",
                        );
                    }
                }
            }
            Ok(None) => {
                if manual {
                    set_main_title(
                        &app,
                        "CRM Operacional AmPm • Você está na versão mais recente",
                    );
                }
            }
            Err(err) => {
                eprintln!("Falha ao verificar atualização: {err}");
                if manual {
                    set_main_title(
                        &app,
                        "CRM Operacional AmPm • Não foi possível verificar atualizações",
                    );
                }
            }
        }
    });
}

fn setup_application(app: &mut tauri::App) -> Result<(), Box<dyn std::error::Error>> {
    // Menu nativo
    let sistema = SubmenuBuilder::new(app, "Sistema")
        .text("open_main", "Abrir CRM")
        .text("reload", "Recarregar CRM")
        .separator()
        .text("quit", "Sair")
        .build()?;

    let exibir = SubmenuBuilder::new(app, "Exibir")
        .text("fullscreen", "Alternar tela cheia")
        .text("minimize", "Minimizar")
        .build()?;

    let ajuda = SubmenuBuilder::new(app, "Ajuda")
        .text("check_update", "Verificar atualizações")
        .separator()
        .text("about", "Sobre o CRM")
        .build()?;

    let app_menu = MenuBuilder::new(app)
        .items(&[&sistema, &exibir, &ajuda])
        .build()?;

    app.set_menu(app_menu)?;

    app.on_menu_event(|app, event| {
        match event.id().as_ref() {
            "open_main" => show_main(app),
            "reload" => reload_crm(app),
            "fullscreen" => toggle_fullscreen(app),
            "minimize" => {
                if let Some(window) = app.get_webview_window("main") {
                    let _ = window.minimize();
                }
            }
            "check_update" => check_for_updates(app.clone(), true),
            "about" => show_about(app),
            "quit" => app.exit(0),
            _ => {}
        }
    });

    // Menu da bandeja do sistema
    let tray_open = MenuItem::with_id(app, "tray_open", "Abrir CRM", true, None::<&str>)?;
    let tray_reload = MenuItem::with_id(app, "tray_reload", "Recarregar CRM", true, None::<&str>)?;
    let tray_update = MenuItem::with_id(
        app,
        "tray_update",
        "Verificar atualizações",
        true,
        None::<&str>,
    )?;
    let tray_about = MenuItem::with_id(app, "tray_about", "Sobre", true, None::<&str>)?;
    let tray_quit = MenuItem::with_id(app, "tray_quit", "Sair", true, None::<&str>)?;
    let tray_menu = Menu::with_items(
        app,
        &[&tray_open, &tray_reload, &tray_update, &tray_about, &tray_quit],
    )?;

    let mut tray_builder = TrayIconBuilder::new()
        .menu(&tray_menu)
        .show_menu_on_left_click(true)
        .tooltip("CRM Operacional AmPm")
        .on_menu_event(|app, event| {
            match event.id.as_ref() {
                "tray_open" => show_main(app),
                "tray_reload" => reload_crm(app),
                "tray_update" => check_for_updates(app.clone(), true),
                "tray_about" => show_about(app),
                "tray_quit" => app.exit(0),
                _ => {}
            }
        });

    if let Some(icon) = app.default_window_icon() {
        tray_builder = tray_builder.icon(icon.clone());
    }

    let _tray = tray_builder.build(app)?;

    // Janela principal remota: sem capacidades Tauri extras.
    let remote_url = CRM_URL.parse().expect("URL do CRM inválida");
    let splash = app.get_webview_window("splashscreen");

    WebviewWindowBuilder::new(
        app,
        "main",
        WebviewUrl::External(remote_url),
    )
    .title("CRM Operacional AmPm")
    .initialization_script(CRM_RESILIENCE_SCRIPT)
    .inner_size(1440.0, 900.0)
    .min_inner_size(1100.0, 700.0)
    .center()
    .resizable(true)
    .maximized(true)
    .visible(false)
    .on_page_load(move |window, payload| {
        if matches!(payload.event(), PageLoadEvent::Finished) {
            let _ = window.show();
            let _ = window.set_focus();

            if let Some(splash_window) = splash.as_ref() {
                let _ = splash_window.close();
            }
        }
    })
    .build()?;

    // Em builds release, verifica silenciosamente por atualização na inicialização.
    // Em `npm run dev`, a checagem automática fica desativada.
    check_for_updates(app.handle().clone(), false);

    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let builder = tauri::Builder::default()
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_single_instance::init(|app, _args, _cwd| {
            show_main(app);
        }))
        .setup(setup_application);

    builder
        .run(tauri::generate_context!())
        .expect("erro ao executar o CRM Operacional AmPm");
}
