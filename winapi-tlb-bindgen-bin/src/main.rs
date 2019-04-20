#![warn(rust_2018_idioms, warnings)]
#![deny(clippy::all, clippy::pedantic)]

#[derive(structopt::StructOpt)]
struct Options {
	#[structopt(help = "path of typelib")]
	filename: std::path::PathBuf,

	#[structopt(long = "emit-dispinterfaces", help = "emit code for DISPINTERFACEs (experimental)")]
	emit_dispinterfaces: bool,
}

fn main() {
	use std::io::Write;

	let Options {
		filename,
		emit_dispinterfaces,
	} = structopt::StructOpt::from_args();

	let build_result = {
		let stdout = std::io::stdout();
		winapi_tlb_bindgen::build(&filename, emit_dispinterfaces, stdout.lock()).unwrap()
	};

	if build_result.num_missing_types > 0 {
		writeln!(&mut std::io::stderr(), "{} referenced types could not be found and were replaced with `__missing_type__`", build_result.num_missing_types).unwrap();
	}

	if build_result.num_types_not_found > 0 {
		writeln!(&mut std::io::stderr(), "{} types could not be found", build_result.num_types_not_found).unwrap();
	}

	for skipped_dispinterface in build_result.skipped_dispinterfaces {
		writeln!(&mut std::io::stderr(), "Dispinterface {} was skipped because --emit-dispinterfaces was not specified", skipped_dispinterface).unwrap();
	}

	for skipped_dispinterface in build_result.skipped_dispinterface_of_dual_interfaces {
		writeln!(&mut std::io::stderr(), "Dispinterface half of dual interface {} was skipped", skipped_dispinterface).unwrap();
	}
}
