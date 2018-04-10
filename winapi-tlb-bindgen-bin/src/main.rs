#[macro_use]
extern crate clap;
extern crate winapi_tlb_bindgen;

fn main() {
	use ::std::io::Write;

	let app = clap_app! {
		@app (app_from_crate!())
		(@arg filename: +required index(1) "filename")
		(@arg ("enable-dispinterfaces"): --("enable-dispinterfaces") "emit code for DISPINTERFACEs (experimental)")
	};

	let matches = app.get_matches();
	let filename = matches.value_of_os("filename").unwrap();
	let filename = std::path::Path::new(filename);
	let emit_dispinterfaces = matches.is_present("enable-dispinterfaces");

	let build_result = {
		let stdout = std::io::stdout();
		let build_result = winapi_tlb_bindgen::build(filename, emit_dispinterfaces, stdout.lock()).unwrap();
		build_result
	};

	if build_result.num_missing_types > 0 {
		writeln!(&mut ::std::io::stderr(), "{} referenced types could not be found and were replaced with `__missing_type__`", build_result.num_missing_types).unwrap();
	}

	if build_result.num_types_not_found > 0 {
		writeln!(&mut ::std::io::stderr(), "{} types could not be found", build_result.num_types_not_found).unwrap();
	}

	for skipped_dispinterface in build_result.skipped_dispinterfaces {
		writeln!(&mut ::std::io::stderr(), "Dispinterface {} was skipped because --emit-dispinterfaces was not specified", skipped_dispinterface).unwrap();
	}

	for skipped_dispinterface in build_result.skipped_dispinterface_of_dual_interfaces {
		writeln!(&mut ::std::io::stderr(), "Dispinterface half of dual interface {} was skipped", skipped_dispinterface).unwrap();
	}
}
