mod common;
mod application;

use windows::core::{IUnknown, Interface, BSTR, GUID, PCWSTR, VARIANT};
use windows::Win32::System::Com::{CoCreateInstance, CoInitialize, GetErrorInfo, IDispatch, IErrorInfo, ITypeInfo, CLSCTX};

use anyhow::{bail, Context, Error, Result};

use common::variant::{EvilVariant, TypedVariant, VariantError};
use common::dispatch::{DispatchError, HasDispatch, Invocation};

use application::Outlook;

use once_cell::sync::OnceCell;

static CO_INITIALIZED : OnceCell<()> = OnceCell::new();


const OBJECT_CONTEXT: CLSCTX = windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
const LOCALE_USER_DEFAULT: u32 = 0x0400;

fn main() -> Result<(), WinError> {
    let outlook = Outlook::new()?;

    

    let folder_path = vec!["ulrik.h@tryg.no", "Inbox"];
    let target_path = vec!["ulrik.h@tryg.no"];

    let folder1_path = vec!["ulrik.h@tryg.no", "Testfolder"];
    let folder2_path = vec!["ulrik.h@tryg.no", "Test 2"];

    let Some(inbox) = outlook.get_folder(folder_path)? else {
        return Ok(())    };
    let Some(target_folder) = outlook.get_folder(target_path)? else {
        return Ok(());    };
    let Some(folder1) = outlook.get_folder(folder1_path)? else {
        return Ok(());    };
    let Some(folder2) = outlook.get_folder(folder2_path)? else {
        return Ok(());    };

    
    let email = inbox.iter()?.next().unwrap();

    let error_info = unsafe {
        GetErrorInfo(0)
    };
    let result = email.move_to(folder1);

    dbg!(result);

    
    match error_info {
        Err(_) => (),
        Ok(info) => unsafe {
            println!(
                "{:?}\n{:?}\n{:?}\n{:?}\n",
                info.GetDescription().unwrap().to_string(),
                info.GetHelpContext().unwrap().to_string(),
                info.GetSource().unwrap().to_string(),
                info.GetHelpFile().unwrap().to_string(),
            )
        }
    };
    
    

    Ok(())
}


pub struct DispatchObject(pub IDispatch);

impl TryFrom<VARIANT> for DispatchObject {
    type Error = Error;
    
    fn try_from(variant: VARIANT) -> std::prelude::v1::Result<Self, Self::Error> {
        let public = EvilVariant::from(variant);
        if public.vt != 9 {
            bail!("VARIANT not dispatch, vt : {}", public.vt);
        }
        let dispatch_ref : &IDispatch = unsafe { std::mem::transmute(&public.union) };
        Ok(DispatchObject(dispatch_ref.clone()))
    }
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

fn wide(rstr : &str) -> PCWSTR {
    let utf16 : Vec<u16> = rstr.encode_utf16().chain(std::iter::once(0)).collect();
    let wide_str : PCWSTR = PCWSTR(utf16.as_ptr());
    wide_str
}

fn bstr(rstr : &str) -> Result<BSTR> {
    let utf16 : Vec<u16> = rstr.encode_utf16().chain(std::iter::once(0)).collect();
    let bstr = BSTR::from_wide(&utf16)?;
    Ok(bstr)
}



impl HasDispatch for DispatchObject {
    fn dispatch(&self) -> &IDispatch {
        self.0.dispatch()
    }
}



#[derive(Debug)]
pub enum WinError {
    VariantError(VariantError),
    DispatchError(DispatchError),
    Internal(windows::core::Error)
}