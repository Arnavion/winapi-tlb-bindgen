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
	let Options {
		filename,
		emit_dispinterfaces,
	} = structopt::StructOpt::from_args();

	let build_result = {
		let stdout = std::io::stdout();
		winapi_tlb_bindgen::build(&filename, emit_dispinterfaces, stdout.lock()).unwrap()
	};

	if build_result.num_missing_types > 0 {
		eprintln!("{} referenced types could not be found and were replaced with `__missing_type__`", build_result.num_missing_types);
	}

	if build_result.num_types_not_found > 0 {
		eprintln!("{} types could not be found", build_result.num_types_not_found);
	}

	for skipped_dispinterface in build_result.skipped_dispinterfaces {
		eprintln!("Dispinterface {} was skipped because --emit-dispinterfaces was not specified", skipped_dispinterface);
	}

	for skipped_dispinterface in build_result.skipped_dispinterface_of_dual_interfaces {
		eprintln!("Dispinterface half of dual interface {} was skipped", skipped_dispinterface);
	}
}
