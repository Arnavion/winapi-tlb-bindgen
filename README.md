Generates bindings to COM interfaces / enums in the style of [winapi-rs v0.3](https://github.com/retep998/winapi-rs/tree/dev)

1. Run against a .tlb

	```powershell
	cargo run -- 'C:\Program Files (x86)\Windows Kits\8.1\Lib\winv6.3\um\x64\MsXml.Tlb' > ~\Desktop\msxml.rs
	if ($LASTEXITCODE -eq 0) { cat ~\Desktop\msxml.rs }
	```

1. Copy the output to winapi and build.

1. ???

1. Profit
