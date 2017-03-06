Generates bindings to COM interfaces / enums in the style of [winapi-rs v0.2](https://github.com/retep998/winapi-rs/tree/0.2)

1. Compile `main.cpp` to `typelib.exe`

1. Run against a .tlb

	```powershell
	.\x64\Debug\typelib.exe 'C:\Program Files (x86)\Windows Kits\8.1\Lib\winv6.3\um\x64\MsXml.Tlb' > ~\Desktop\msxml.rs
	if ($LASTEXITCODE -eq 0) { cat ~\Desktop\msxml.rs }
	```

1. Copy the output to winapi and build.

1. ???

1. Profit

---

Requires a new macro `UNION2!` in winapi that transmutes `self` instead of `self.$field`. The existing `UNION!` macro acts on the field of a struct `S` that is of a type that's a union `U`, thus it generates impls of `S` that transmute the field into one of the union variants. This codegen instead generates a newtype for the union and expects the `UNION2!` macro to generate impls for the newtype that transmute itself to the variants.
