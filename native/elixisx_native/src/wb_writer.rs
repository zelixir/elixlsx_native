use error::{ExcelError, ExcelResult};
use rustler::dynamic::get_type;
use rustler::types::{Binary, ListIterator, MapIterator, OwnedBinary};
use rustler::{Decoder, Encoder, Env, Term, TermType};
use std::cmp::Eq;
use std::io::{BufWriter, Cursor, Seek, SeekFrom, Write};
use wb_compiler::{SheetCompInfo, WorkbookCompInfo};
use workbook::{Sheet, Workbook};
use xml_writer::XmlWriter;
use zip::result::{ZipError, ZipResult};
use zip::write::FileOptions;
use zip::ZipWriter;

pub fn create_excel<'a, W: Write + Seek>(
    workbook: Workbook<'a>,
    mut wci: WorkbookCompInfo,
    writer: W,
) -> ExcelResult<()> {
    let options = FileOptions::default()
        .compression_method(::zip::CompressionMethod::Deflated)
        .unix_permissions(0o755);

    let mut writer = ExcelWriter(ZipWriter::new(writer), options);
    writer.write_doc_props_dir(&workbook)?;
    writer.write_rels_dir()?;
    writer.write_xl_dir(&workbook, &mut wci)?;

    Ok(())
}
pub fn create_excel_data<'a>(
    workbook: Workbook<'a>,
    wci: WorkbookCompInfo,
) -> ExcelResult<Vec<u8>> {
    let mut cursor = Cursor::new(vec![]);
    create_excel(
        workbook,
        wci,
        ::std::io::BufWriter::new(SeekableVec(&mut cursor)),
    )?;
    Ok(cursor.get_ref().to_vec())
}

struct ExcelWriter<W: Write + Seek>(ZipWriter<W>, FileOptions);

impl<W: Write + Seek> ExcelWriter<W> {
    fn write_doc_props_dir(&mut self, workbook: &Workbook) -> ZipResult<()> {
        // app.xml
        self.0.start_file("docProps/app.xml", self.1)?;
        self.0
            .write(::xml_templates::doc_props_app("1.00".to_string()).as_bytes())?;
        // core.xml
        self.0.start_file("docProps/core.xml", self.1)?;
        self.0.write(
            ::xml_templates::doc_props_core(workbook.datetime.clone(), None, None).as_bytes(),
        )?;
        // self.write_xml()
        Ok(())
    }

    fn write_rels_dir(&mut self) -> ZipResult<()> {
        self.0.start_file("_rels/.rels", self.1)?;
        self.0.write(::xml_templates::rels_dotrels().as_bytes())?;

        Ok(())
    }

    fn write_xl_dir(&mut self, workbook: &Workbook, wci: &mut WorkbookCompInfo) -> ExcelResult<()> {
        self.write_xl_worksheets_dir(workbook, wci)?;
        ::xml_templates::write_xl_styles(&mut self.0, wci)?;
        Ok(())
    }

    fn write_xl_worksheets_dir(
        &mut self,
        workbook: &Workbook,
        wci: &mut WorkbookCompInfo,
    ) -> ExcelResult<()> {
        for (sheet, filename) in workbook.sheets.iter().zip(
            wci.sheet_info
                .iter()
                .map(|x| x.filename.clone())
                .collect::<Vec<String>>(),
        ) {
            self.0
                .start_file(format!("xl/worksheets/{}", filename), self.1)?;
            ::xml_templates::write_sheet(&mut self.0, sheet, wci)?;
        }

        Ok(())
    }
}

struct SeekableVec<'a>(&'a mut Cursor<Vec<u8>>);

impl<'a> Seek for SeekableVec<'a> {
    fn seek(&mut self, pos: SeekFrom) -> ::std::io::Result<u64> {
        let pos = match pos {
            SeekFrom::Current(0) => self.0.position(),
            SeekFrom::Start(pos) => pos,
            _ => Err(::std::io::Error::from(::std::io::ErrorKind::Other))?,
        };
        self.0.set_position(pos);
        Ok(pos)
    }
}

impl<'a> Write for SeekableVec<'a> {
    fn write(&mut self, buf: &[u8]) -> ::std::io::Result<usize> {
        self.0.write(buf)
    }
    fn flush(&mut self) -> ::std::io::Result<()> {
        Ok(())
    }
}
