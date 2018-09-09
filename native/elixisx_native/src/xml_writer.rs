use error::ExcelResult;
use regex::Regex;
use std::io::{Result, Write};

pub trait XmlWriter {
    fn write_string(&mut self, string: &ToString) -> Result<usize>;

    fn write_xml<F>(
        &mut self,
        tag: &ToString,
        attrs: Vec<(&ToString, &ToString)>,
        inner: F,
    ) -> ExcelResult<()>
    where
        F: FnOnce(&mut Self) -> ExcelResult<()>,
    {
        self.write_string(&"<")?;
        self.write_string(tag)?;
        self.write_attrs(attrs)?;
        self.write_string(&">")?;

        inner(self)?;

        self.write_string(&"</")?;
        self.write_string(tag)?;
        self.write_string(&">")?;

        Ok(())
    }
    fn write_xml_empty_tag(
        &mut self,
        tag: &ToString,
        attrs: Vec<(&ToString, &ToString)>,
    ) -> ExcelResult<()> {
        self.write_string(&"<")?;
        self.write_string(tag)?;
        self.write_attrs(attrs)?;
        self.write_string(&"/>")?;
        Ok(())
    }

    fn write_attrs(&mut self, attrs: Vec<(&ToString, &ToString)>) -> ExcelResult<()> {
        for (name, value) in attrs {
            self.write_string(&" ")?;
            self.write_string(name)?;
            self.write_string(&"=\"")?;
            self.write_string(value)?;
            self.write_string(&"\"")?;
        }
        Ok(())
    }
}

impl<T: Write> XmlWriter for T {
    fn write_string(&mut self, string: &ToString) -> Result<usize> {
        self.write(string.to_string().as_bytes())
    }
}

pub struct Escaped<'a>(pub &'a ToString);

impl<'a> ToString for Escaped<'a> {
    fn to_string(&self) -> String {
        xml_escape(self.0.to_string())
    }
}

pub fn xml_escape(input: String) -> String {
    lazy_static! {
        static ref REGEX: Regex = Regex::new("[<>&\"']").unwrap();
    }
    {
        if let Some(first) = REGEX.find(&input) {
            let first = first.start();
            let len = input.len();
            let mut output: Vec<u8> = Vec::with_capacity(len + len / 2);
            output.extend_from_slice(input[0..first].as_bytes());
            let rest = input[first..].bytes();
            for c in rest {
                match c {
                    b'<' => output.extend_from_slice(b"&lt;"),
                    b'>' => output.extend_from_slice(b"&gt;"),
                    b'&' => output.extend_from_slice(b"&amp;"),
                    b'\'' => output.extend_from_slice(b"&apos;"),
                    b'"' => output.extend_from_slice(b"&quot;"),
                    _ => output.push(c),
                }
            }
            Some(unsafe { String::from_utf8_unchecked(output) })
        } else {
            None
        }
    }.unwrap_or(input)
}
