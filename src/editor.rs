use eframe::egui;

use crate::json_builder;

pub struct Editor {
    input_file_path: Option<String>,
    export_msg: Option<String>,
    main_sheet_name: String,
}

impl Editor {
    pub fn new(_cc: &eframe::CreationContext<'_>) -> Self {
        Self { 
            input_file_path: None,
            export_msg: None,
            main_sheet_name: "main".into(),
        }
    }

    fn draw(&mut self, ui: &mut egui::Ui) {
        ui.heading("Excel to Json");
        ui.horizontal(|ui| {
            ui.label("excel path: ");
            
            if ui.button("Open file…").clicked() {
                if let Some(path) = rfd::FileDialog::new().pick_file() {
                    self.input_file_path = Some(path.display().to_string());
                }
            }
        });

        if let Some(picked_path) = &self.input_file_path {
            ui.horizontal(|ui| {
                ui.label("Picked file:");
                ui.monospace(picked_path);
            });
        }

        ui.horizontal(|ui| {
            ui.text_edit_singleline(&mut self.main_sheet_name);

            if ui.button("Export").clicked() {
                if let Some(path) = rfd::FileDialog::new().save_file() {
                    self.export_msg = Some(format!("{:?}", path));
                    if let Ok(_) = json_builder::build(self.input_file_path.as_ref(), &path, &self.main_sheet_name) {
                        self.export_msg = Some("export succ".into());
                    } else {
                        self.export_msg = Some("export fail".into());
                    }
                }
            }
        });

        if let Some(msg) = self.export_msg.as_ref() {
            ui.separator();
            ui.horizontal(|ui| {
                ui.label("export msg:");
                ui.monospace(msg);
            });
        }
    }
}

impl eframe::App for Editor {
    fn update(&mut self, ctx: &eframe::egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            self.draw(ui);
        });
    }
}