fn main() {
	let cargo = std::env::var_os("CARGO").unwrap();

	let msxml = {
		let source_dir = std::env::var_os("CARGO_MANIFEST_DIR").unwrap();
		let source_dir: std::path::PathBuf = source_dir.into();
		let bindgen_dir = source_dir.parent().unwrap();
		let bindgen_output =
			std::process::Command::new(cargo)
			.args(&["run", "--", r"C:\Program Files (x86)\Windows Kits\10\Lib\10.0.16299.0\um\x64\MsXml.Tlb"])
			.current_dir(bindgen_dir)
			.output().unwrap();
		assert!(bindgen_output.status.success());
		bindgen_output.stdout
	};

	let mut msxml_rs = {
		let msxml_rs = std::env::var_os("OUT_DIR").unwrap();
		let mut msxml_rs: std::path::PathBuf = msxml_rs.into();
		msxml_rs.push("msxml.rs");
		let msxml_rs = std::fs::OpenOptions::new().create(true).write(true).truncate(true).open(msxml_rs).unwrap();
		std::io::BufWriter::new(msxml_rs)
	};

	std::io::Write::write_all(&mut msxml_rs, &msxml).unwrap();
}
