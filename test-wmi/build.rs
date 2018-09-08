extern crate winapi_tlb_bindgen;

fn main() {
	let out_dir: std::path::PathBuf = std::env::var_os("OUT_DIR").unwrap().into();

	let midl_command_status =
		std::process::Command::new("midl.exe") // Expected to be running in "x64 Native Tools Command Prompt"
		.arg(r"C:\Program Files (x86)\Windows Kits\10\Include\10.0.17134.0\um\WbemCli.Idl")
		.arg("/tlb")
		.arg("WbemCli.tlb")
		.current_dir(&out_dir)
		.status().unwrap();
	assert!(midl_command_status.success());

	let wbemcli_rs = {
		let wbemcli_rs = out_dir.join("wbemcli.rs");
		let wbemcli_rs = std::fs::OpenOptions::new().create(true).write(true).truncate(true).open(wbemcli_rs).unwrap();
		std::io::BufWriter::new(wbemcli_rs)
	};

	let _ =
		winapi_tlb_bindgen::build(
			&out_dir.join("WbemCli.tlb"),
			false,
			wbemcli_rs,
		).unwrap();
}
