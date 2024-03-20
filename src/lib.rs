mod common;
mod application;

use common::{dispatch::DispatchError, variant::VariantError};
use once_cell::sync::OnceCell;
use windows::{core::{BSTR, PCWSTR}, Win32::System::Com::{CoInitialize, CLSCTX, EXCEPINFO}};

const OBJECT_CONTEXT: CLSCTX = windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
const LOCALE_USER_DEFAULT: u32 = 0x0400;

static CO_INITIALIZED : OnceCell<()> = OnceCell::new();

#[derive(Debug)]
enum WinError {
    VariantError(VariantError),
    DispatchError(DispatchError),
    Internal(windows::core::Error)
}


fn wide(rstr : &str) -> PCWSTR {
    let utf16 : Vec<u16> = rstr.encode_utf16().chain(std::iter::once(0)).collect();
    let wide_str : PCWSTR = PCWSTR(utf16.as_ptr());
    wide_str
}

fn bstr(rstr : &str) -> Result<BSTR, WinError> {
    let utf16 : Vec<u16> = rstr.encode_utf16().chain(std::iter::once(0)).collect();
    BSTR::from_wide(&utf16).map_err(
        |e| WinError::Internal(e)
    )
}

fn co_initialize() {
    CO_INITIALIZED.get_or_init(
        || {
            let Ok(_) = (unsafe {
                CoInitialize(None).ok()
            }) else {
                panic!("Failed to CoInitialize");
            };
        }
    );
}