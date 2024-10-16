use serde_json::Map;
use umya_spreadsheet::{Worksheet, Spreadsheet};

struct JsonBuilder<'a> {
    sheet_builder: SheetBuilder<'a>,
}

impl<'a> JsonBuilder<'a> {
    pub fn build(book: &'a Spreadsheet, entry_sheet_name: &str) -> Self {
        Self {
            // book: book,
            sheet_builder: SheetBuilder::build(entry_sheet_name, book),
        }
    }

    pub fn json(&self, path: &std::path::Path) {
        
        let d = self.sheet_builder.json();
        // to_string_pretty
        // to_string
        if d.len() == 1 {
            std::fs::write(path, serde_json::to_string(&d[0]).unwrap()).expect("export: json write fail");
        } else if d.len() > 1 {
            std::fs::write(path, serde_json::to_string(&d).unwrap()).expect("export: json write fail");
        } else {
            println!("nothing~");
        }
    }
}

struct SheetBuilder<'a> {
    col_type_list: Vec<String>,
    col_name_list: Vec<String>,
    book: &'a Spreadsheet,
    sheet: &'a Worksheet,
}

impl<'a> SheetBuilder<'a> {
    pub fn build(name: &str, book: &'a Spreadsheet) -> Self {
        println!("sheet build, name: {}", name);
        let sheet = book.get_sheet_by_name(name).expect(&format!("sheet not found: {}", name));
        Self {
            sheet: sheet,
            book: book,
            col_type_list: SheetBuilder::build_col_type(sheet),
            col_name_list: SheetBuilder::build_col_name(sheet),
        }
    }

    pub fn json(&self) -> Vec<serde_json::Value> {
        println!("types: {:?}", self.col_type_list);
        println!("names: {:?}", self.col_name_list);
        self.gen_row()
    }

    fn build_col_type(sheet: &'a Worksheet) -> Vec<String> {
        let col_len = sheet.get_collection_by_row(&1).len() as u32;
        let mut list = vec![];
        for j in 1 ..= col_len {
            if let Some(cell) = sheet.get_cell((j, 2)) {
                list.push(format!("{}", cell.get_value()));
            }
        }
        list
    }

    fn build_col_name(sheet: &'a Worksheet) -> Vec<String> {
        let col_len = sheet.get_collection_by_row(&1).len() as u32;
        let mut list = vec![];
        for j in 1 ..= col_len {
            if let Some(cell) = sheet.get_cell((j, 3)) {
                list.push(format!("{}", cell.get_value()));
            }
        }
        list
    }

    pub fn gen_row(&self) -> Vec<serde_json::Value> {
        let row_len = self.sheet.get_collection_by_column(&1).len() as u32;
        let col_len = self.sheet.get_collection_by_row(&1).len() as u32;
        println!("cols: {}, rows: {}", col_len, row_len);
        let mut list: Vec<serde_json::Value> = vec![];
        for i in 4 ..= row_len {
            let mut map = Map::new();
            for j in 1 ..= col_len {
                if let Some(cell) = self.sheet.get_cell((j, i)) {
                    if let Some(field_val) = self.gen_cell(j, &format!("{}", cell.get_value())) {
                        let field_name = self.col_name_list.get((j-1) as usize).unwrap();
                        map.insert(field_name.clone(), field_val);
                    }
                } else {
                    break
                }
            }
            list.push(map.into());
        }
        return list;
    }

    fn gen_cell(&self, index: u32, txt: &String) -> Option<serde_json::Value> {
        let col_type = self.col_type_list.get((index - 1) as usize).unwrap();
        match col_type as &str {
            "bool" => SheetBuilder::bool_2_v(txt),
            "number" => SheetBuilder::int_2_v(txt),
            "string" => SheetBuilder::str_2_v(txt),
            "[]bool" => SheetBuilder::arr_bool_2_v(txt),
            "[]number" => SheetBuilder::arr_num_2_v(txt),
            "[]string" => SheetBuilder::arr_str_2_v(txt),
            "[]object" => SheetBuilder::arr_obj_2_v(self, txt),
            _ => {
                panic!("数据类型不支持~, {}", col_type);
            }
        }
    }

    fn bool_2_v(txt: &str) -> Option<serde_json::Value> {
        let new_txt = &txt.trim().to_lowercase();
        match new_txt as &str {
            "true" => Some(true.into()),
            "false" => Some(false.into()),
            _ => panic!("bool 转换失败, 值 不为 true 或 false"),
        }
    } 

    fn int_2_v(txt: &str) -> Option<serde_json::Value> {
        let new_txt = txt.trim();
        if let Ok(v) = new_txt.parse::<i32>() {
            return Some(v.into());
        } else {
            let v = new_txt.parse::<f32>().expect(&format!("convert to int fail~, raw txt: {}.", new_txt));
            return Some(v.into());
        }
    }

    fn str_2_v(txt: &str) -> Option<serde_json::Value> {
        let txt = txt.trim();
        Some(txt.into())
    }

    fn arr_str_2_v(txt: &str) -> Option<serde_json::Value> {
        let txt = txt.trim();
        let list = txt.split(",");
        let mut tartget_list = vec![];

        for item in list.into_iter() {
            tartget_list.push( SheetBuilder::str_2_v(item) );
        }
        Some(tartget_list.into())
    }

    fn arr_bool_2_v(txt: &str) -> Option<serde_json::Value> {
        let txt = txt.trim();
        let txt = txt.replace("[", "");
        let txt = txt.replace("]", "");
        let list = txt.split(",");
        let mut tartget_list = vec![];

        for item in list.into_iter() {
            tartget_list.push( SheetBuilder::bool_2_v(item) );
        }
        Some(tartget_list.into())
    }

    fn arr_num_2_v(txt: &str) -> Option<serde_json::Value> {
        let txt = txt.trim();
        let txt = txt.replace("[", "");
        let txt = txt.replace("]", "");
        let list = txt.split(",");
        let mut tartget_list = vec![];

        for item in list.into_iter() {
            tartget_list.push( SheetBuilder::int_2_v(item) );
        }
        Some(tartget_list.into())
    }

    fn arr_obj_2_v(&self, txt: &str) -> Option<serde_json::Value> {
        let txt = txt.trim();
        let txt = txt.replace("[", "");
        let txt = txt.replace("]", "");
        let mut list = txt.split(",");
        if let Some(method) = list.next() {
            let method = method.trim();
            match method {
                "#include" => {
                    let sheet_name = list.next();
                    println!("#include {:?}", sheet_name);
                    return self.include_sheet(sheet_name);
                }
                _ => {
                    println!("method dont't support~, method: {}", method);
                }
            }
        }
        None
    }

    fn include_sheet(&self, sheet_name: Option<&str>) -> Option<serde_json::Value> {
        if let Some(name) = sheet_name {
            let name = name.trim();
            return Some(SheetBuilder::build(name, self.book).json().into());
        };
        None
    }
}

pub fn build(xlsx_path: Option<&String>, output_path: &std::path::PathBuf, entry_sheet_name: &str) -> Result<(), String> {
    let xlsx_path = xlsx_path.unwrap();
    let path = std::path::Path::new(&xlsx_path);
    let book = umya_spreadsheet::reader::xlsx::read(path).expect("read xlsx fail");
    
    JsonBuilder::build(&book, entry_sheet_name).json(output_path);
    Ok(())
}