#[macro_use]
extern crate rustler;
#[macro_use]
extern crate lazy_static;
extern crate regex;
// extern crate rustler_codegen;
extern crate chrono;
extern crate zip;
use rustler::types::OwnedBinary;
use rustler::{Env, Error, NifResult, Term};

mod error;
mod xml_writer;
mod wb_compiler;
mod wb_writer;
mod workbook;
mod xml_templates;
mod util;

rustler_export_nifs! {
    "Elixir.Elixlsx.Native",
    [
     ("write_excel", 1, write_excel)
    ],
    None
}

fn write_excel<'a>(env: Env<'a>, args: &[Term<'a>]) -> NifResult<Term<'a>> {
    let (workbook, wci) = wb_compiler::make_workbook_comp_info(env, args)?;
    let data = wb_writer::create_excel_data(workbook, wci)?;
    let mut bin = OwnedBinary::new(data.len()).ok_or(Error::Atom("write_bin_error"))?;
    bin.as_mut_slice().clone_from_slice(data.as_slice());
    Ok(bin.release(env).to_term(env))
}


