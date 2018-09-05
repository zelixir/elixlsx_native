use error::ExcelResult;
use rustler::dynamic::get_type;
use rustler::types::{Binary, ListIterator, MapIterator, OwnedBinary};
use rustler::{Decoder, Env, Error, NifResult, Term, TermType};
use std::cmp::Eq;
use std::collections::HashMap;
use std::hash::Hash;
use workbook::{Sheet, Workbook};

pub fn make_workbook_comp_info<'a>(
    env: Env<'a>,
    args: &[Term<'a>],
) -> ExcelResult<(Workbook<'a>, WorkbookCompInfo)> {
    let workbook: Workbook = args[0].decode()?;

    let (sci, next_rid) = make_sheet_info(&workbook.sheets, 2);

    let wci = WorkbookCompInfo {
        sheet_info: sci,
        next_free_xl_rid: next_rid,
        ..Default::default()
    };
    // try!(wci.compinfo_from_sheets(&workbook.sheets));
    // try!(wci.regist_all_cell_style());

    Ok((workbook, wci))
}

fn make_sheet_info(sheets: &Vec<Sheet>, first_free_rid: i32) -> (Vec<SheetCompInfo>, i32) {
    let len = sheets.len() as i32;
    let li = (0..len)
        .map(|x| SheetCompInfo::make(x + 1, x + first_free_rid))
        .collect();
    (li, first_free_rid + len)
}

#[derive(Default)]
pub struct SheetCompInfo {
    pub rId: String,
    pub filename: String,
    pub sheetId: i32,
}

impl SheetCompInfo {
    fn make(idx: i32, rid: i32) -> SheetCompInfo {
        SheetCompInfo {
            rId: format!("rId{}", rid),
            filename: format!("sheet{}.xml", idx),
            sheetId: idx,
        }
    }
}

#[derive(Default)]
pub struct WorkbookCompInfo {
    pub sheet_info: Vec<SheetCompInfo>,
    pub stringdb: DB<String>,
    pub fontdb: DB<Font>,
    pub filldb: DB<String>,
    pub cellstyledb: DB<CellStyle>,
    pub numfmtdb: DB<String>,
    pub borderstyledb: DB<BorderStyle>,
    pub next_free_xl_rid: i32,
}

impl WorkbookCompInfo {
    // fn compinfo_from_sheets(&mut self, sheets: &Vec<Sheet>) -> NifResult<()> {
    //     for sheet in sheets {
    //         try!(self.compinfo_from_rows(&sheet.rows));
    //     }
    //     Ok(())
    // }
    // fn compinfo_from_rows<'a>(&mut self, rows: &Term<'a>) -> NifResult<()> {
    //     let list: ListIterator = try!(rows.decode());
    //     for row in list {
    //         for cell in try!(row.decode::<ListIterator>()) {
    //             self.compinfo_cell_pass(cell);
    //         }
    //     }
    //     Ok(())
    // }
    // fn compinfo_cell_pass<'a>(&mut self, cell: Term<'a>) {
    //     match get_type(cell) {
    //         TermType::List => {
    //             let list: ListIterator = try!(cell.decode());
    //             list.next().map(|x| self.compinfo_cell_pass(x));
    //             self.compinfo_cell_pass_style(list.collect());
    //         }
    //         TermType::Binary => cell
    //         _ => (),
    //     };
    // }
    // fn compinfo_cell_pass_style<'a>(&mut self, props: Vec<Term<'a>>) {}
    // fn regist_all_cell_style(&mut self) -> NifResult<()> {
    //     Ok(())
    // }
}

#[derive(Default)]
pub struct DB<T: Eq + Hash> {
    pub data: HashMap<T, i32>,
}

impl<T: Eq + Hash + Clone> DB<T> {
    pub fn get_id(&mut self, key: &T) -> i32 {
        match self.data.get(key) {
            Some(&id) => id,
            _ => {
                let id = self.data.len() as i32 + 1;
                self.data.insert(key.clone(), id);
                id
            }
        }
    }
}

#[derive(Default, Eq, PartialEq, Hash, Clone)]
pub struct Border {
    pub type_: String,
    pub style: String,
    pub color: String,
}
impl<'a> Border {
    fn new(map: &HashMap<String, Term<'a>>, type_: String) -> NifResult<Self> {
        Ok(Border {
            type_: type_,
            style: get_keyword_value(map, "style", Default::default())?,
            color: get_keyword_value(map, "color", Default::default())?,
        })
    }
}

#[derive(Default, Eq, PartialEq, Hash, Clone)]
pub struct BorderStyle {
    pub left: Border,
    pub right: Border,
    pub top: Border,
    pub bottom: Border,
    pub diagonal: Border,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}
impl<'a> BorderStyle {
    fn new(map: &HashMap<String, Term<'a>>) -> NifResult<Self> {
        fn get_border<'a>(map: &HashMap<String, Term<'a>>, name: &str) -> NifResult<Border> {
            let li: ListIterator = map.get(name).ok_or_else(|| Error::BadArg)?.decode()?;
            let map = ::workbook::decode_keyword_list(li)?;
            Border::new(&map, name.to_string())
        }

        Ok(BorderStyle {
            left: get_border(map, "left")?,
            right: get_border(map, "right")?,
            top: get_border(map, "top")?,
            bottom: get_border(map, "bottom")?,
            diagonal: get_border(map, "diagonal")?,
            diagonal_down: get_bool(map, "diagonal_down"),
            diagonal_up: get_bool(map, "diagonal_up"),
        })
    }
}

#[derive(Default, Eq, PartialEq, Hash, Clone)]
pub struct Font {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub size: i32,
    pub color: String,
    pub wrap_text: bool,
    pub align_horizontal: String,
    pub align_vertical: String,
    pub font: String,
}

impl<'a> Font {
    fn new(map: &HashMap<String, Term<'a>>) -> NifResult<Self> {
        Ok(Font {
            bold: get_bool(map, "bold"),
            italic: get_bool(map, "italic"),
            underline: get_bool(map, "underline"),
            strike: get_bool(map, "strike"),
            size: map.get("size").map_or(Err(Error::BadArg), |x| x.decode())?,
            color: get_keyword_value(map, "color", Default::default())?,
            wrap_text: get_bool(map, "wrap_text"),
            align_horizontal: get_keyword_value(map, "align_horizontal", Default::default())?,
            align_vertical: get_keyword_value(map, "align_vertical", Default::default())?,
            font: get_keyword_value(map, "font", Default::default())?,
        })
    }
}

#[derive(Default, Eq, PartialEq, Hash, Clone)]
pub struct CellStyle {
    pub font: Font,
    pub fill: String,
    pub numfmt: String,
    pub border: BorderStyle,
}

impl<'a> CellStyle {
    pub fn new(list: ListIterator<'a>) -> NifResult<Self> {
        let map = ::workbook::decode_keyword_list(list)?;
        Ok(CellStyle {
            font: Font::new(&map)?,
            fill: get_keyword_value(&map, "bg_color", Default::default())?,
            numfmt: get_numfmt(&map)?,
            border: BorderStyle::new(&map)?,
        })
    }
    pub fn is_date(&self) -> bool {
        self.numfmt.contains("yy")
    }
}

fn get_numfmt<'a>(map: &HashMap<String, Term<'a>>) -> NifResult<String> {
    Ok(if map.contains_key("yyyymmdd") {
        "yyyy-mm-dd".to_string()
    } else if map.contains_key("datetime") {
        "yyyy-mm-dd h:mm:ss".to_string()
    } else if let Some(num_format) = map.get("num_format") {
        num_format.decode()?
    } else {
        "".to_string()
    })
}

fn get_keyword_value<'a, T: Default + Decoder<'a>>(
    map: &HashMap<String, Term<'a>>,
    key: &str,
    default: T,
) -> NifResult<T> {
    map.get(key).map_or(Ok(default), |term| term.decode())
}

fn get_bool<'a>(map: &HashMap<String, Term<'a>>, key: &str) -> bool {
    map.get(key)
        .map_or(false, |x| x.decode::<bool>().unwrap_or(false))
}
