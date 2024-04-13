mod common;
mod application;

use windows::core::{w, IUnknown, Interface, BSTR, GUID, HRESULT, PCWSTR, VARIANT};
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, GetErrorInfo, IDispatch, CLSCTX, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO};

use anyhow::{bail, Error, Result};

use common::variant::{opt_out_arg, EvilVariant, TypedVariant, VariantError};
use common::dispatch::{DispatchError, HasDispatch, Invocation};

use application::{Folder, Outlook};

use once_cell::sync::OnceCell;

static CO_INITIALIZED : OnceCell<()> = OnceCell::new();


const OBJECT_CONTEXT: CLSCTX = windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
const LOCALE_USER_DEFAULT: u32 = 0x0400;

fn main() -> Result<(), WinError> {
    co_initialize();

    let outlook = Outlook::new()?;
    let folder_path = vec!["Ulrik.Hansen@student.uib.no", "Inbox"];
    let target_path = vec!["Ulrik.Hansen@student.uib.no", "Inbox", "test 2"];

    let Some(inbox) = outlook.get_folder(folder_path)? else {
        return Ok(())    };
    let Some(target_folder) = outlook.get_folder(target_path)? else {
        return Ok(());    };

    let TypedVariant::Dispatch(picked_folder) = outlook.session().call("PickFolder", Invocation::Method, None)? else {
        panic!("pickfolder failed")
    };


    let TypedVariant::Bstr(storeID1) = picked_folder.prop("StoreID")? else {
        panic!()
    };

    let TypedVariant::Bstr(entryID1) = picked_folder.prop("EntryID")? else {
        panic!()
    };

    let TypedVariant::Bstr(storeID2) = target_folder.prop("StoreID")? else {
        panic!()
    };

    let TypedVariant::Bstr(entryID2) = target_folder.prop("EntryID")? else {
        panic!()
    };
    
    println!("{}", entryID1);

    println!("{}", entryID2);


    let args = DISPPARAMS {
        cArgs : 2,
        rgvarg : vec![VARIANT::from(storeID2), VARIANT::from(entryID2)].as_mut_ptr(),
        ..Default::default()
    };

    let TypedVariant::Dispatch(test_from_id) = outlook.session().call("GetFolderFromID", Invocation::Method, Some(args))? else {
        panic!("ID DIDNT WORK")
    };

    let email = inbox.iter()?.next().unwrap();

    email.move_to(Folder(test_from_id))?;
    

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
                CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()
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