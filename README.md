Generates bindings to COM interfaces / enums in the style of [winapi-rs v0.2](https://github.com/retep998/winapi-rs/tree/0.2)

1. Compile `main.cpp` to `winapi-tlb-bindgen.exe`

1. Run against a .tlb

	```powershell
	# Output compatible with winapi v0.2
	.\winapi-tlb-bindgen.exe 'C:\Program Files (x86)\Windows Kits\8.1\Lib\winv6.3\um\x64\MsXml.Tlb' > ~\Desktop\msxml.rs
	if ($LASTEXITCODE -eq 0) { cat ~\Desktop\msxml.rs }

	or

	```powershell
	# Output compatible with winapi v0.3
	.\winapi-tlb-bindgen.exe 0.3 'C:\Program Files (x86)\Windows Kits\8.1\Lib\winv6.3\um\x64\MsXml.Tlb' > ~\Desktop\msxml.rs
	if ($LASTEXITCODE -eq 0) { cat ~\Desktop\msxml.rs }
	```

1. Copy the output to winapi and build.

1. ???

1. Profit

---

Requires the `UNION2!` winapi macro that was added to v0.3. Some changes to the existing macros are also needed. The winapi-0.2.patch in this repository can be applied to a clone of winapi-rs's `0.2` branch to get everything working.
