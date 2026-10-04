use tauri::{
    AppHandle, Manager, WebviewUrl,
    menu::{Menu, MenuBuilder, MenuItem, SubmenuBuilder},
    tray::TrayIconBuilder,
    webview::{PageLoadEvent, WebviewWindowBuilder},
};

use tauri_plugin_updater::UpdaterExt;

const CRM_URL: &str = "https://email-campaign-19.preview.emergentagent.com";

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

    // A janela principal usa exatamente o mesmo frontend React publicado no ambiente web.
    let remote_url = CRM_URL.parse().expect("URL do CRM inválida");
    let splash = app.get_webview_window("splashscreen");

    WebviewWindowBuilder::new(
        app,
        "main",
        WebviewUrl::External(remote_url),
    )
    .title("CRM Operacional AmPm")
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
