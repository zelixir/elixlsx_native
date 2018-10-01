use error::ExcelResult;
use std::collections::HashMap;
use std::io::Write;
use wb_compiler::WorkbookCompInfo;
use workbook::Workbook;

pub fn create_excel<'a>(
    workbook: Workbook<'a>,
    mut wci: WorkbookCompInfo,
) -> ExcelResult<HashMap<String, Vec<u8>>> {
    let mut writer = ExcelWriter::new();
    writer.write_doc_props_dir(&workbook)?;
    writer.write_rels_dir()?;
    writer.write_xl_dir(&workbook, &mut wci)?;

    ::xml_templates::write_content_types(
        &mut writer.start_file(&"[Content_Types].xml"),
        &wci.sheet_info,
    )?;
    Ok(writer.data)
}

#[derive(Default)]
struct ExcelWriter {
    data: HashMap<String, Vec<u8>>,
}

impl ExcelWriter {
    fn new() -> Self {
        Default::default()
    }
    fn start_file(&mut self, filename: &ToString) -> &mut Write {
        let buf: Vec<u8> = Vec::with_capacity(4096);
        self.data.insert(filename.to_string(), buf);
        self.data.get_mut(&filename.to_string()).unwrap()
    }

    fn write_doc_props_dir(&mut self, workbook: &Workbook) -> ::std::io::Result<&mut Self> {
        // app.xml
        self.start_file(&"docProps/app.xml")
            .write(::xml_templates::doc_props_app("1.00".to_string()).as_bytes())?;
        // core.xml
        self.start_file(&"docProps/core.xml").write(
            ::xml_templates::doc_props_core(workbook.datetime.clone(), None, None).as_bytes(),
        )?;
        Ok(self)
    }

    fn write_rels_dir(&mut self) -> ::std::io::Result<()> {
        self.start_file(&"_rels/.rels")
            .write(::xml_templates::rels_dotrels().as_bytes())?;
        Ok(())
    }

    fn write_xl_dir(&mut self, workbook: &Workbook, wci: &mut WorkbookCompInfo) -> ExcelResult<()> {
        self.write_xl_worksheets_dir(workbook, wci)?;
        ::xml_templates::write_xl_styles(&mut self.start_file(&"xl/styles.xml"), wci)?;
        ::xml_templates::wite_string_db(
            &mut self.start_file(&"xl/sharedStrings.xml"),
            &wci.stringdb,
        )?;
        ::xml_templates::write_workbook_xml(
            &mut self.start_file(&"xl/workbook.xml"),
            &workbook.sheets,
            &wci.sheet_info,
        )?;
        ::xml_templates::write_xl_rels(
            &mut self.start_file(&"xl/_rels/workbook.xml.rels"),
            &wci.sheet_info,
            wci.next_free_xl_rid,
        )?;
        Ok(())
    }

    fn write_xl_worksheets_dir(
        &mut self,
        workbook: &Workbook,
        wci: &mut WorkbookCompInfo,
    ) -> ExcelResult<()> {
        for (filename, sheet) in wci.sheet_info
            .iter()
            .map(|x| x.filename.clone())
            .collect::<Vec<String>>()
            .iter()
            .zip(workbook.sheets.iter())
        {
            let mut writer = self.start_file(&format!("xl/worksheets/{}", filename));
            ::xml_templates::write_sheet(&mut writer, sheet, wci)?;
        }
        Ok(())
    }
}
