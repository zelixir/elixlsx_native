use rustler::dynamic::{get_type, TermType};
use rustler::types::{ListIterator, MapIterator};
use rustler::{Decoder, Error};
use rustler::{NifResult, Term};
use std::cmp::Eq;
use std::collections::HashMap;
use std::hash::Hash;

pub struct Workbook<'a> {
    pub sheets: Vec<Sheet<'a>>,
    pub datetime: String,
}

impl<'a> Decoder<'a> for Workbook<'a> {
    fn decode(term: Term<'a>) -> NifResult<Self> {
        let mut wb = Workbook {
            sheets: vec![],
            datetime: format!("{:?}", ::chrono::Utc::now()),
        };
        let map = (to_map(term))?;
        if let Some(sheets) = map.get("sheets") {
            let sheets: ListIterator = (sheets.decode())?;
            wb.sheets = try!(sheets.map(|x| x.decode::<Sheet>()).collect());
        }
        if let Some(datetime) = map.get("datetime") {
            match get_type(*datetime) {
                TermType::Number => {
                    wb.datetime = format!(
                        "{:?}",
                        ::chrono::NaiveDateTime::from_timestamp(datetime.decode::<i64>()?, 0),
                    )
                }
                TermType::Binary => wb.datetime = datetime.decode::<String>()?,
                _ => (),
            }
        }

        Ok(wb)
    }
}

pub struct Sheet<'a> {
    pub name: String,
    pub rows: Term<'a>,
    pub col_widths: HashMap<i32, i32>,
    pub row_heights: HashMap<i32, i32>,
    pub merge_cells: Vec<(String, String)>,
    pub pane_freeze: Option<(i32, i32)>,
    pub show_grid_lines: bool,
}

impl<'a> Decoder<'a> for Sheet<'a> {
    fn decode(term: Term<'a>) -> NifResult<Self> {
        let map = (to_map(term))?;
        let re = Sheet {
            name: map.get("name")
                .and_then(|x| x.decode().ok())
                .unwrap_or("".to_string()),
            rows: map.get("rows").map(|&x| x).ok_or(Error::BadArg)?,
            col_widths: to_int_map(map.get("col_widths"))?,
            row_heights: to_int_map(map.get("row_heights"))?,
            merge_cells: decode_merge_cells(map.get("merge_cells"))?,
            pane_freeze: map.get("pane_freeze").and_then(|x| x.decode().ok()),
            show_grid_lines: map.get("show_grid_lines")
                .and_then(|&x| x.atom_to_string().ok())
                .map_or(false, |x| x == "true"),
        };
        Ok(re)
    }
}

fn to_map<'a>(term: Term<'a>) -> NifResult<HashMap<String, Term<'a>>> {
    let re = term.decode::<MapIterator>()?
        .filter_map(|(k, v)| match k.atom_to_string() {
            Ok(k) => Some((k, v)),
            _ => None,
        })
        .collect();
    Ok(re)
}

fn to_int_map<'a>(term: Option<&Term<'a>>) -> NifResult<HashMap<i32, i32>> {
    match term {
        Some(term) => decode_hash_map(*term),
        _ => Ok(Default::default()),
    }
}

fn decode_merge_cells<'a>(term: Option<&Term<'a>>) -> NifResult<Vec<(String, String)>> {
    match term {
        Some(term) => {
            let li: ListIterator = term.decode()?;
            li.map(|x| x.decode::<(String, String)>()).collect()
        }
        _ => Ok(Default::default()),
    }
}

pub fn decode_hash_map<'a, K: Eq + Hash + Decoder<'a>, V: Decoder<'a>>(
    term: Term<'a>,
) -> NifResult<HashMap<K, V>> {
    term.decode::<MapIterator>()?
        .map(|(k, v)| Ok((k.decode::<K>()?, v.decode::<V>()?)))
        .collect()
}

pub fn decode_keyword_list<'a>(list: ListIterator<'a>) -> NifResult<HashMap<String, Term<'a>>> {
    list.map(|x| {
        let x = ::rustler::types::tuple::get_tuple(x)?;
        if x.len() == 2 {
            Ok((x[0].decode::<String>()?, x[1]))
        } else {
            Err(Error::BadArg)
        }
    }).collect()
}

pub enum CellValue {
    ExcelTS(String),
    Formula(String, HashMap<String, String>),
    String(String),
    Number(String),
    Date(::chrono::NaiveDateTime),
    Empty,
    None,
}

impl<'a> CellValue {
    pub fn new(term: Term<'a>, is_date: bool) -> NifResult<Self> {
        Ok(match (get_type(term), is_date) {
            (TermType::Tuple, true) => {
                let ((y, m, d), (h, mm, s)) = term.decode::<((i32, u32, u32), (u32, u32, u32))>()?;
                CellValue::Date(::chrono::NaiveDate::from_ymd(y, m, d).and_hms(h, mm, s))
            }
            (TermType::Tuple, false) => {
                let li = ::rustler::types::tuple::get_tuple(term)?;
                if li.len() >= 2 && li.len() <= 3 {
                    let t: &str = li[0].decode()?;
                    match t {
                        "excelts" => CellValue::ExcelTS(li[1].decode()?),
                        "formula" => {
                            let formula: String = li[1].decode()?;
                            let mut opts: HashMap<String, String> = HashMap::new();
                            if li.len() == 3 {
                                opts = decode_hash_map(li[2])?;
                            }
                            CellValue::Formula(formula, opts)
                        }
                        _ => CellValue::None,
                    }
                } else {
                    CellValue::None
                }
            }
            (TermType::Number, _) => CellValue::Number(term.decode::<String>()?),
            (TermType::Binary, _) => CellValue::String(term.decode::<String>()?),
            (TermType::Atom, _) => {
                if term.decode::<String>()? == "empty" {
                    CellValue::Empty
                } else {
                    CellValue::None
                }
            }
            //nil or others
            _ => CellValue::None,
        })
    }
}
