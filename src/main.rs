// #![cfg_attr(not(debug_assertions), windows_subsystem = "windows")] // hide console window on windows in release

pub mod json_builder;
pub mod parse_args;

// use clap::Parser;
// use parse_args::{parse_args, Args};

pub mod editor;

use eframe::egui;

fn main() {
//     let Args {
//         input_path,
//         output_path,
//         entry_sheet,
//     } = parse_args();

    //let path = std::path::Path::new(&output_path);
    // let book = umya_spreadsheet::reader::xlsx::read(path).expect("read xlsx fail");

    // match json_builder::build(
    //     Some(&input_path),
    //     &path.to_path_buf(),
    //     &entry_sheet.unwrap(),
    // ) {
    //     Err(e) => {
    //         println!("build config fail, {}", e);
    //     }
    //     Ok(_) => {
    //         println!("build config succ~");
    //     }
    // }

    let options = eframe::NativeOptions {
        initial_window_size: Some(egui::vec2(500_f32, 300_f32)),
        ..Default::default()
    };

    if let Err(e) = eframe::run_native("Excel2Json", options, Box::new(|cc| Box::new(editor::Editor::new(cc)))) {
        println!("run err: {:?}", e);
    }
}
