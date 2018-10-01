#[macro_use]
extern crate rustler;
#[macro_use]
extern crate lazy_static;
extern crate chrono;
extern crate regex;
use rustler::{Encoder, Env, NifResult, Term};

mod error;
mod util;
mod wb_compiler;
mod wb_writer;
mod workbook;
mod xml_templates;
mod xml_writer;

rustler_export_nifs! {
    "Elixir.Elixlsx.Native",
    [
     ("write_excel_nif", 1, write_excel)
    ],
    None
}

fn write_excel<'a>(env: Env<'a>, args: &[Term<'a>]) -> NifResult<Term<'a>> {
    let (workbook, wci) = wb_compiler::make_workbook_comp_info(args)?;
    Ok(wb_writer::create_excel(workbook, wci)?
        .into_iter()
        .map(|(k, v)| (k.into_bytes(), unsafe { String::from_utf8_unchecked(v) }))
        .collect::<Vec<(Vec<u8>, String)>>()
        .encode(env))
}
