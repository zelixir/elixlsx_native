use error::ExcelResult;
use rustler::dynamic::get_type;
use rustler::types::ListIterator;
use rustler::{Term, TermType};
use std::collections::HashMap;
use util::to_excel_coords;
use wb_compiler::{Border, BorderStyle, CellStyle, Font, SheetCompInfo, WorkbookCompInfo, DB};
use workbook::{CellValue, Sheet};
use xml_writer::{Escaped, XmlWriter};

pub fn write_content_types<T: XmlWriter>(
  writer: &mut T,
  scis: &Vec<SheetCompInfo>,
) -> ExcelResult<()> {
  writer.write_string(&r###"<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    "###)?;
  for sci in scis {
    writer.write_string(&format!(r###"
        <Override PartName="/xl/worksheets/{}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    "###, sci.filename))?;
  }
  writer.write_string(&r###"
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    </Types>
    "###)?;
  Ok(())
}
pub fn write_xl_rels<T: XmlWriter>(
  writer: &mut T,
  scis: &Vec<SheetCompInfo>,
  next_free_xl_rid: i32,
) -> ExcelResult<()> {
  writer.write_string(&r#"<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  "#)?;
  for sci in scis {
    writer.write_string(&format!("<Relationship Id=\"{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/{}\"/>", sci.rid, sci.filename))?;
  }
  writer.write_string(&format!(r#"
        <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      </Relationships>
  "#, next_free_xl_rid))?;
  Ok(())
}
pub fn write_workbook_xml<T: XmlWriter>(
  writer: &mut T,
  sheets: &Vec<Sheet>,
  scis: &Vec<SheetCompInfo>,
) -> ExcelResult<()> {
  writer.write_string(&r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="Calc"/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      "#)?;
  for (sheet, sci) in sheets.iter().zip(scis) {
    writer.write_xml_empty_tag(
      &"sheet",
      vec![
        (&"name", &Escaped(&sheet.name)),
        (&"sheetId", &sci.sheet_id),
        (&"state", &"visible"),
        (&"r:id", &sci.rid),
      ],
    )?;
  }
  writer.write_string(&r#"
    </sheets>
    <calcPr fullCalcOnLoad="1" iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
    </workbook>
      "#)?;

  Ok(())
}
pub fn wite_string_db<T: XmlWriter>(writer: &mut T, stringdb: &DB<String>) -> ExcelResult<()> {
  let list = stringdb.sorted_list();
  let len = list.len();
  writer.write_string(&format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">
"#, len, len))?;

  for (string, _) in list {
    writer.write_string(&"<si><t>")?;
    writer.write_string(&Escaped(string))?;
    writer.write_string(&"</t></si>")?;
  }
  writer.write_string(&"</sst>")?;
  Ok(())
}
pub fn write_xl_styles<T: XmlWriter>(
  writer: &mut T,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<()> {
  writer.write_string(&r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#)?;
  writer.write_string(
    &r#"      <cellStyleXfs count="1">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" applyAlignment="1">
          <alignment wrapText="1"/>
        </xf>
      </cellStyleXfs>
  "#,
  )?;
  let cell_styles: Vec<(&CellStyle, &i32)> = wci.cellstyledb.sorted_list();
  writer.write_string(&format!(
    r#"      <cellXfs count="{}">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" xfId="0"/>"#,
    &(1 + cell_styles.len())
  ))?;
  for (style, _) in cell_styles {
    write_cell_style(
      writer,
      style,
      &mut wci.fontdb,
      &mut wci.filldb,
      &mut wci.numfmtdb,
      &mut wci.borderstyledb,
    )?;
    writer.write_string(&"\n")?;
  }
  writer.write_string(&"</cellXfs>")?;
  write_numfmts(writer, wci.numfmtdb.sorted_list())?;

  let font_list = wci.fontdb.sorted_list();
  writer.write_xml(&"fonts", vec![(&"count", &(1 + font_list.len()))], |w| {
    w.write_string(&"<font />")?;
    for (font, _) in font_list {
      write_font(w, font)?;
      w.write_string(&"\n")?;
    }
    Ok(())
  })?;

  let fill_list = wci.filldb.sorted_list();
  writer.write_xml(&"fills", vec![(&"count", &(2 + fill_list.len()))], |w| {
    w.write_string(
      &r#"
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    "#,
    )?;
    for (fill, _) in fill_list {
      write_fill(w, fill)?;
      w.write_string(&"\n")?;
    }
    Ok(())
  })?;

  let border_list = wci.borderstyledb.sorted_list();
  writer.write_xml(
    &"borders",
    vec![(&"count", &(1 + border_list.len()))],
    |w| {
      w.write_string(&r#"<border />"#)?;
      for (border, _) in border_list {
        write_border_style(w, border)?;
        w.write_string(&"\n")?;
      }
      Ok(())
    },
  )?;
  writer.write_string(&"</styleSheet>")?;

  Ok(())
}
fn write_border<T: XmlWriter>(writer: &mut T, border: &Border) -> ExcelResult<()> {
  let style = match border.style.as_str() {
    "dash_dot" => "dashDot",
    "dash_dot_dot" => "dashDotDot",
    any => any,
  }.to_string();
  let mut attrs: Vec<(&ToString, &ToString)> = vec![];
  if style != "" {
    attrs.push((&"style", &style));
  }
  writer.write_xml(&border.type_, attrs, |w| {
    if border.color != "" {
      w.write_string(&format!(
        "<color rgb=\"{}\" />",
        to_argb_color(border.color.clone())
      ))?;
    }
    Ok(())
  })
}
fn write_border_style<T: XmlWriter>(writer: &mut T, border: &BorderStyle) -> ExcelResult<()> {
  writer.write_xml(
    &"border",
    vec![
      (&"diagonalUp", &border.diagonal_up),
      (&"diagonalDown", &border.diagonal_down),
    ],
    |w| {
      write_border(w, &border.left)?;
      write_border(w, &border.right)?;
      write_border(w, &border.top)?;
      write_border(w, &border.bottom)?;
      if border.diagonal_down || border.diagonal_up {
        write_border(w, &border.diagonal)?;
      } else {
        w.write_string(&"<diagonal></diagonal>")?;
      }
      Ok(())
    },
  )
}

fn write_fill<T: XmlWriter>(writer: &mut T, fill: &String) -> ExcelResult<()> {
  writer.write_xml(&"fill", vec![], |w| {
    if fill != "" {
      w.write_string(&format!(
        "<patternFill patternType=\"solid\"><fgColor rgb=\"{}\" /></patternFill>",
        to_argb_color(fill.clone())
      ))?;
    }
    Ok(())
  })
}
fn write_font<T: XmlWriter>(writer: &mut T, font: &Font) -> ExcelResult<()> {
  writer.write_xml(&"font", vec![], |w| {
    // TODO: Add more underline properties, see here:
    // http://webapp.docx4java.org/OnlineDemo/ecma376/SpreadsheetML/ST_UnderlineValues.html
    if font.bold {
      w.write_string(&"<b val=\"1\"/>")?;
    }
    if font.italic {
      w.write_string(&"<i val=\"1\"/>")?;
    }
    if font.underline {
      w.write_string(&"<u val=\"single\"/>")?;
    }
    if font.strike {
      w.write_string(&"<strike val=\"1\"/>")?;
    }
    if font.size > 0 {
      w.write_string(&format!("<sz val=\"{}\"/>", font.size))?;
    }
    if font.color != "" {
      w.write_string(&format!(
        "<color rgb=\"{}\" />",
        to_argb_color(font.color.clone())
      ))?;
    }
    if font.font != "" {
      w.write_string(&format!("<name val=\"{}\" />", font.font))?;
    }

    Ok(())
  })
}

fn write_cell_style<T: XmlWriter>(
  writer: &mut T,
  style: &CellStyle,
  fontdb: &mut DB<Font>,
  filldb: &mut DB<String>,
  numfmtdb: &mut DB<String>,
  borderstyledb: &mut DB<BorderStyle>,
) -> ExcelResult<()> {
  let font_id = if let Some(font) = &style.font {
    fontdb.get_id(&font) + 1
  } else {
    0
  };
  let fill_id = if style.fill != "" {
    filldb.get_id(&style.fill) + 2
  } else {
    0
  };

  let numfmt_id = if style.numfmt != "" {
    numfmtdb.get_id(&style.numfmt) + 164
  } else {
    0
  };
  let border_id = borderstyledb.get_id(&style.border);
  let alignment_attrs = style
    .font
    .as_ref()
    .map_or(vec![], |x| x.get_alignment_attributes());

  let mut style_attrs: Vec<(&ToString, &ToString)> = vec![
    (&"borderId", &border_id),
    (&"fillId", &fill_id),
    (&"fontId", &font_id),
    (&"numFmtId", &numfmt_id),
    (&"xfId", &0),
  ];
  if alignment_attrs.len() > 0 {
    style_attrs.push((&"applyAlignment", &1));
  }

  writer.write_xml(&"xf", style_attrs, |w| {
    if alignment_attrs.len() > 0 {
      w.write_xml_empty_tag(&"alignment", alignment_attrs)?;
    }
    Ok(())
  })?;

  Ok(())
}
fn write_numfmts<T: XmlWriter>(writer: &mut T, numfmts: Vec<(&String, &i32)>) -> ExcelResult<()> {
  let len = numfmts.len();
  if len > 0 {
    writer.write_xml(&"numFmts", vec![(&"count", &len)], |w| {
      for (fmt, index) in numfmts {
        w.write_string(&format!(
          r#"<numFmt numFmtId="{}" formatCode="{}" />\n"#,
          index + 164,
          fmt
        ))?;
      }
      Ok(())
    })?;
  }
  Ok(())
}

pub fn write_sheet<T: XmlWriter>(
  writer: &mut T,
  sheet: &Sheet,
  wci: &mut WorkbookCompInfo,
) -> ExcelResult<()> {
  writer.write_string(&r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
  writer.write_string(&r#"</sheetData>"#)?;
  write_merge_cells(writer, &sheet.merge_cells)?;
  writer.write_string(
    &r##"
      <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
    </worksheet>
  "##,
  )?;
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
  if sheet.col_widths.len() > 0 {
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
    let (content, style_id) = split_into_content_style(cell, wci)?;
    let r = to_excel_coords(row_index, i);
    match content {
      CellValue::String(string) => {
        let id = wci.stringdb.get_id(&string);
        writer.write_string(&format!(
          r##"<c r="{}" s="{}" t="s">
              <v>{}</v>
              </c>"##,
          r, style_id, id
        ))?;
      }
      CellValue::Empty => {
        writer.write_string(&format!(r##"<c r="{}" s="{}"></c>"##, r, style_id))?;
      }
      CellValue::Number(num) => {
        writer.write_string(&format!(
          r##"<c r="{}" s="{}" t="n">
              <v>{}</v>
              </c>"##,
          r, style_id, num
        ))?;
      }
      CellValue::Formula(formular, opts) => {
        let value = match opts.get("value") {
          Some(value) => format!("<v>{}</v>", value),
          _ => "".to_string(),
        };
        writer.write_string(&format!(
          r##"<c r="{}"
              s="{}">
              <f>{}</f>
              {}
              </c>"##,
          r, style_id, formular, value
        ))?;
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
) -> ExcelResult<(CellValue, i32)> {
  Ok(match get_type(cell) {
    TermType::List => {
      let mut li: ListIterator = cell.decode()?;
      match li.next() {
        Some(term) => {
          let cell_style = CellStyle::new(li)?;
          let cell_value = CellValue::new(term, cell_style.is_date())?;
          (cell_value, wci.cellstyledb.get_id(&cell_style) + 1)
        }
        _ => (CellValue::None, 0),
      }
    }
    _ => (CellValue::new(cell, false)?, 0),
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
fn write_merge_cells<T: XmlWriter>(
  writer: &mut T,
  merge_cells: &Vec<(String, String)>,
) -> ExcelResult<()> {
  if merge_cells.len() > 0 {
    writer.write_xml(&"mergeCells", vec![(&"count", &merge_cells.len())], |w| {
      for (from, to) in merge_cells {
        w.write_string(&format!("<mergeCell ref=\"{}:{}\"/>", from, to))?;
      }
      Ok(())
    })?;
  }
  Ok(())
}

pub fn to_argb_color(color: String) -> String {
  if color != "" {
    format!("#{}", color[1..].to_string())
  } else {
    "".to_string()
  }
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
  format!( r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
  <dc:language>{}</dc:language>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
  <cp:revision>{}</cp:revision>
</cp:coreProperties>
"#, time, language.unwrap_or("en-US".to_string()), time, revision.unwrap_or(1))
}

pub fn rels_dotrels() -> String {
  r#"<?xml version="1.0" encoding="UTF-8"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  </Relationships>
"#.to_string()
}
