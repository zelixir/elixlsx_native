use rustler::dynamic::get_type;
use rustler::types::{Binary, ListIterator, MapIterator, OwnedBinary};
use rustler::{Decoder, Env, Term, TermType};
use std::cmp::Eq;
use std::collections::HashMap;
use std::hash::Hash;
use workbook::{Sheet, Workbook};
use error::{ExcelResult};

pub fn make_workbook_comp_info<'a>(env: Env<'a>, args: &[Term<'a>]) -> ExcelResult<(Workbook<'a>, WorkbookCompInfo)> {
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
    pub count: i32,
}

#[derive(Default, Eq, PartialEq, Hash)]
pub struct Border {
    pub type_: String,
    pub style: String,
    pub coloe: String,
}

#[derive(Default, Eq, PartialEq, Hash)]
pub struct BorderStyle {
    pub left: Border,
    pub right: Border,
    pub top: Border,
    pub bottom: Border,
    pub diagonal: Border,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}

#[derive(Default, Eq, PartialEq, Hash)]
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
#[derive(Default, Eq, PartialEq, Hash)]
pub struct CellStyle {
    pub font: Font,
    pub fill: String,
    pub numfmt: String,
    pub border: BorderStyle,
}




