use error::ExcelResult;
use rustler::dynamic::get_type;
use rustler::types::{Binary, ListIterator, MapIterator, OwnedBinary};
use rustler::{Decoder, Encoder, Env, Term, TermType};
use std::collections::HashMap;
use util::to_excel_coords;
use wb_compiler::{CellStyle, Font, SheetCompInfo, WorkbookCompInfo};
use workbook::{CellValue, Sheet};
use xml_writer::XmlWriter;

fn write_sheet_rows<T: XmlWriter>(
  writer: &mut T,
  sheet: &Sheet,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<()> {
  let mut i = 1;

  let rows: ListIterator = sheet.rows.decode()?;
  for r in rows {
    writer.write_xml(&"row", get_row_attr(&i, &sheet.row_heights), |w| {
      write_sheet_cols(w, &r, i, wci)
    })?;
    i = i + 1;
  }
  Ok(())
}

fn write_sheet_cols<'a, T: XmlWriter>(
  writer: &mut T,
  row: &Term<'a>,
  row_index: i32,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<()> {
  let mut i = 1;

  let cols: ListIterator = row.decode()?;
  for cell in cols {
    let (content, style_id, style) = split_into_content_style(cell, wci)?;
    let r = to_excel_coords(row_index, i);
    match content {
      CellValue::String(string) => {
        let id = wci.stringdb.get_id(&string);
        writer.write_string(format!(
          r##"<c r="{}" s="{}" t="s">
              <v>{}</v>
              </c>"##,
          r, style_id, string
        ))?;
      }
      CellValue::Empty => {
        writer.write_string(format!(r##"<c r="{}" s="{}"></c>"##, r, style_id))?;
      }
      CellValue::ExcelTS(num) =>{
        writer.write_string(format!(
          r##"<c r="{}" s="{}" t="n">
              <v>{}</v>
              </c>"##,
          r, style_id, num
        ))?;
      CellValue::Formula(formular, opts) =>{

      }
      CellValue::Number(num) =>{

      }
      CellValue::Date(num) =>{

      }
      _ => (),
    }
    i = i + 1;
  }
  Ok(())
}

fn split_into_content_style<'a>(
  cell: Term<'a>,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<(CellValue, i32, Option<CellStyle>)> {
  Ok(match get_type(cell) {
    TermType::List => {
      let mut li: ListIterator = cell.decode()?;
      match li.next() {
        Some(term) => {
          let cell_style = CellStyle::new(li)?;
          let cell_value = CellValue::new(term, cell_style.is_date())?;
          (
            cell_value,
            wci.cellstyledb.get_id(&cell_style),
            Some(cell_style),
          )
        }
        _ => (CellValue::None, 0, None),
      }
    }
    _ => (CellValue::new(cell, false)?, 0, None),
  })
}

fn get_row_attr<'a>(
  row_index: &'a i32,
  row_heights: &'a HashMap<i32, i32>,
) -> Vec<(&'a ToString, &'a ToString)> {
  let mut re: Vec<(&'a ToString, &'a ToString)> = vec![(&"r", row_index)];
  if let Some(height) = row_heights.get(&row_index) {
    re.push((&"customHeight", &"1"));
    re.push((&"customHeight", height));
  }
  re
}

pub fn doc_props_app(ver: String) -> String {
  format!( r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Elixlsx</Application>
  <AppVersion>{}</AppVersion>
</Properties>
"#, ver)
}

pub fn doc_props_core(time: String, language: Option<String>, revision: Option<i32>) -> String {
  format!( r#"
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
  <dc:language>{}</dc:language>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
  <cp:revision>{}</cp:revision>
</cp:coreProperties>
"#, time, language.unwrap_or("en-US".to_string()), time, revision.unwrap_or(1))
}

pub fn rels_dotrels() -> String {
  r#"
  <?xml version="1.0" encoding="UTF-8"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  </Relationships>
"#.to_string()
}

pub fn write_xl_styles<T: XmlWriter>(writer: &mut T, wci: &WorkbookCompInfo) -> ExcelResult<()> {
  writer.write_string(&r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#.to_string())?;
  write_numfmts(writer, &wci.numfmtdb.data)?;
  Ok(())
}

fn write_numfmts<T: XmlWriter>(writer: &mut T, numfmts: &HashMap<String, i32>) -> ExcelResult<()> {
  let len = numfmts.len();
  if len > 0 {
    writer.write_xml(&"numFmts", vec![(&"count", &len)], |w| Ok(()))?;
  }
  Ok(())
}

pub fn write_sheet<T: XmlWriter>(
  writer: &mut T,
  sheet: &Sheet,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<()> {
  writer.write_string(&r#"
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetPr filterMode="false">
      <pageSetUpPr fitToPage="false"/>
    </sheetPr>
    <dimension ref="A1"/>
    <sheetViews>
    <sheetView workbookViewId="0" 
  "#)?;
  if !sheet.show_grid_lines {
    writer.write_string(&" showGridLines=\"0\" ")?;
  }
  writer.write_string(&">")?;
  writer.write_string(&make_sheet_view(&sheet))?;
  writer.write_string(
    &r#"
      </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="12.8"/>
  "#,
  )?;
  wrtie_col_widths(writer, sheet)?;
  writer.write_string(&r#"<sheetData>"#)?;
  write_sheet_rows(writer, sheet, wci)?;
  writer.write_string(&r#"<sheetData>"#)?;

  Ok(())
}

fn make_sheet_view(sheet: &Sheet) -> String {
  let pane = match sheet.pane_freeze {
    Some((_, 0)) => "bottomLeft",
    Some((0, _)) => "topRight",
    Some((x, y)) if x > 0 && y > 0 => "bottomRight",
    _ => "",
  };

  let (selection_pane_attr, panel_xml) = match sheet.pane_freeze {
    Some((x, y)) if x > 0 || y > 0 => {
      let top_left_cell = ::util::to_excel_coords(x + 1, y + 1);
      let s = format!("pane=\"{}\"", pane);
      let p = format!("<pane xSplit=\"{}\" ySplit=\"{}\" topLeftCell=\"{}\" activePane=\"{}\" state=\"frozen\" />", 
        x, y, top_left_cell, pane);
      (s, p)
    }
    _ => ("".to_string(), "".to_string()),
  };

  format!(
    "{}<selection {} activeCell=\"A1\" sqref=\"A1\" />",
    selection_pane_attr, panel_xml
  )
}
fn wrtie_col_widths<T: XmlWriter>(writer: &mut T, sheet: &Sheet) -> ExcelResult<()> {
  if sheet.col_widths.len() >= 0 {
    writer.write_xml(&"cols", vec![], |w| {
      let mut li = sheet
        .col_widths
        .iter()
        .map(|(&k, &v)| (k, v))
        .collect::<Vec<(i32, i32)>>();
      li.sort_by(|a, b| a.0.cmp(&b.0));
      for (col, width) in li {
        w.write_string(&format!(
          r#"<col min="{}" max="{}" width="{}" customWidth="1" />"#,
          col, col, width
        ))?;
      }
      Ok(())
    })?;
  }
  Ok(())
}
