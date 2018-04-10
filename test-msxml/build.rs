extern crate winapi_tlb_bindgen;

fn main() {
	let msxml_rs = {
		let msxml_rs = std::env::var_os("OUT_DIR").unwrap();
		let mut msxml_rs: std::path::PathBuf = msxml_rs.into();
		msxml_rs.push("msxml.rs");
		let msxml_rs = std::fs::OpenOptions::new().create(true).write(true).truncate(true).open(msxml_rs).unwrap();
		std::io::BufWriter::new(msxml_rs)
	};

	let _ =
		winapi_tlb_bindgen::build(
			std::path::Path::new(r"C:\Program Files (x86)\Windows Kits\10\Lib\10.0.16299.0\um\x64\MsXml.Tlb"),
			false,
			msxml_rs,
		).unwrap();
}
