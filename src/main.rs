mod common;
mod outlook;

use windows::core::{IUnknown, Interface, BSTR, GUID, PCWSTR, VARIANT};
use windows::Win32::System::Com::{CoCreateInstance, CoInitialize, IDispatch, ITypeInfo, CLSCTX, DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, TYPEATTR};

use anyhow::{bail, Error, Result, Context};

use common::variant::{EvilVariant, SafeVariant};

use once_cell::sync::OnceCell;

static CO_INITIALIZED : OnceCell<()> = OnceCell::new();


const OBJECT_CONTEXT: CLSCTX = windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
const LOCALE_USER_DEFAULT: u32 = 0x0400;

fn main() -> Result<()> {
    let outlook = Outlook::new()?;

    let namespace_variant = outlook.get("Session")?;
    let namespace_dispatch = DispatchObject::try_from(namespace_variant)?;

    let folders_variant = namespace_dispatch.get("Folders")?;
    let folders_dispatch = DispatchObject::try_from(folders_variant)?;

    let first_folder_variant = folders_dispatch.call("GetFirst", Invocation::Method, vec![])?;
    let first_folder_dispatch = DispatchObject::try_from(first_folder_variant)?;

    let name_variant = first_folder_dispatch.get("Class")?;
    let name_var = SafeVariant::from(name_variant);

    match name_var {
        SafeVariant::Bstr(string) => {
            println!("{}", string);
        }
        SafeVariant::Dispatch(dispatch) => {
            dbg!(dispatch);
        },
        SafeVariant::Int32(num) => {
            println!("Class id: {}", num);
        }
        _ => panic!("waaaa"),
    };

    Ok(())
}

pub struct Outlook(pub IDispatch);

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

impl Outlook {
    fn new() -> Result<Self> {
        co_initialize();
    
        //let class_id : GUID = unsafe {
        //    windows::Win32::System::Com::CLSIDFromProgID(w!("Outlook.Application")) // Wide string pointer macro
        //}?;
    
        let class_id = GUID::from("0006F03A-0000-0000-C000-000000000046");
        let raw_ptr = &class_id as *const GUID;
    
        let unknown : IUnknown  = unsafe {
            CoCreateInstance(raw_ptr, None, OBJECT_CONTEXT)
        }?;
        
        let dispatch : IDispatch = unknown.cast()?;
        Ok(Outlook(dispatch))
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

// Don't use outside same scope
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

impl HasDispatch for IDispatch {
    fn dispatch(&self) -> &IDispatch {
        self
    }
}

impl HasDispatch for Outlook {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}


#[repr(u16)]
enum Invocation {
    Method = 1,
    PropertyGet = 2,
    PropertySet = 4,
}


trait HasDispatch {
    fn dispatch(&self) -> &IDispatch;

    fn get_dispid(&self, member_name : &str) -> Result<i32> {
        let mut rgdispid : i32 = 0;
        let wide_str = wide(member_name);
    
        let _ = unsafe {
            self.dispatch().GetIDsOfNames(
                &GUID::zeroed(), // Useless param
                &wide_str as *const PCWSTR, // Method name
                1, // # of method names
                LOCALE_USER_DEFAULT, // Localization
                &mut rgdispid as *mut i32 // dispid pointer
            )
    
        }.with_context(|| format!("Failed to find member: {}", member_name))?;
        Ok(rgdispid)
    }

    fn get(&self, property_name : &str) -> Result<VARIANT> {
        self.call(property_name, Invocation::Method, vec![])
    }
     
    fn call(&self, method_name : &str, flag : Invocation, args : Vec<VARIANT>) -> Result<VARIANT> {
        let dispatch = self.dispatch();

        let dispid = self.get_dispid(method_name)?;

        let params = dispparams(args );

        let mut exception : EXCEPINFO = EXCEPINFO::default();

        let mut result = VARIANT::new();
        unsafe {
            if let Err(_) = dispatch.Invoke(
                dispid,
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_FLAGS(flag as u16),
                &params as *const DISPPARAMS,
                Some(&mut result as *mut VARIANT),
                Some(&mut exception as *mut EXCEPINFO),
                None,
            ) {
                bail!(format!("{:?}",exception))
            };
        };

        Ok(result)
    }

    fn get_guid(&self) -> Result<GUID> {
        let type_info : ITypeInfo = unsafe{ self.dispatch().GetTypeInfo(0, LOCALE_USER_DEFAULT) }?;
    
        let attr_ptr = unsafe { type_info.GetTypeAttr()}?;
        if attr_ptr.is_null() {
            bail!("Attribute null")
        }

        let attr : TYPEATTR = unsafe {*attr_ptr };
        Ok(attr.guid)
    }
}


fn dispparams(mut vars : Vec<VARIANT>) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg : vars.as_mut_ptr() as *mut VARIANT,
        rgdispidNamedArgs : std::ptr::null_mut() as *mut i32,
        cArgs : vars.len() as u32,
        cNamedArgs : 0,
    }
}

