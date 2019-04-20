#![allow(non_camel_case_types, non_snake_case, unused)]

use winapi::{ENUM, RIDL, STRUCT};
use winapi::ctypes::c_void;
use winapi::shared::guiddef::GUID;
use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypes::BSTR;
use winapi::um::oaidl::{SAFEARRAY, VARIANT};
use winapi::um::unknwnbase::{IUnknown, IUnknownVtbl, LPUNKNOWN};
use winapi::um::winnt::LPCWSTR;

include!(concat!(env!("OUT_DIR"), "/wbemcli.rs"));
