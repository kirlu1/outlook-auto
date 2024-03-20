use anyhow::{bail, Context, Result};
use windows::{core::{GUID, PCWSTR, VARIANT}, Win32::System::Com::{IDispatch, ITypeInfo, DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, TYPEATTR}};

use crate::{wide, WinError, LOCALE_USER_DEFAULT};

use super::variant::{SafeVariant, VariantError};

#[derive(Debug)]
pub enum DispatchError {
    InvokeError {
        exception : EXCEPINFO,
        invoked_name : String,
    },
    DispidError {
        name : String,
        error : windows::core::Error,
    }
}

#[repr(u16)]
pub enum Invocation {
    Method = 1,
    PropertyGet = 2,
    PropertySet = 4,
}


pub trait HasDispatch {
    fn dispatch(&self) -> &IDispatch;

    fn get_dispid(&self, member_name : &str) -> Result<i32, WinError> {
        let mut rgdispid : i32 = 0;
        let wide_str = wide(member_name);
    
        
        if let Err(e) = unsafe {
            self.dispatch().GetIDsOfNames(
                &GUID::zeroed(), // Useless param
                &wide_str as *const PCWSTR, // Method name
                1, // # of method names
                LOCALE_USER_DEFAULT, // Localization
                &mut rgdispid as *mut i32 // dispid pointer
            )
        } {
            return Err(WinError::DispatchError(DispatchError::DispidError {
                name : member_name.to_string(),
                error : e
            }))
        };

        Ok(rgdispid)
    }

    fn prop(&self, property_name : &str) -> Result<SafeVariant, WinError> {
        self.call(property_name, Invocation::Method, vec![])
    }
     
    fn call(&self, method_name : &str, flag : Invocation, args : Vec<VARIANT>) -> Result<SafeVariant, WinError> {
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
                return Err(WinError::DispatchError(DispatchError::InvokeError { exception, invoked_name: method_name.to_string() }))
            };
        };

        SafeVariant::try_from(result)
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

impl HasDispatch for IDispatch {
    fn dispatch(&self) -> &IDispatch {
        self
    }
}

