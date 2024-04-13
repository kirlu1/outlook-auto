use windows::{
    core::{w, IUnknown, Interface, GUID, VARIANT}, 
    Win32::System::Com::{CLSIDFromProgID, CoCreateInstance, IDispatch, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS}
};

use core::cell::OnceCell;
use std::{os::raw::c_void, sync::{Arc, Mutex, OnceLock}};

use crate::{bstr, co_initialize, common::{dispatch::{HasDispatch, Invocation}, variant::{EvilVariant, TypedVariant, VariantError}}, WinError, LOCALE_USER_DEFAULT, OBJECT_CONTEXT};


pub struct Outlook {
    pub app : IDispatch,
    pub namespace : IDispatch,
}

impl Outlook {
    pub fn new() -> Result<Self, WinError> {
        co_initialize();
    
        let app : IDispatch = unsafe {
            let clsid = CLSIDFromProgID(w!("Outlook.Application")).expect("Couldn't get CLSID");
            CoCreateInstance(&clsid, None, OBJECT_CONTEXT)
        }.map_err(
            |e| WinError::Internal(e)
        )?;

        let params = DISPPARAMS {
            cArgs : 1,
            rgvarg : &mut VARIANT::from("MAPI"),
            ..Default::default()
        };

        let mut namespace_variant = VARIANT::default();
 
        unsafe {
            app.Invoke(
                272,
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &params,
                Some(&mut namespace_variant),
                None,
                None,
            )
        }.unwrap();

        let namespace = IDispatch::try_from(&namespace_variant).expect("couldnt cast VARIANT to IDispatch");

        Ok(Outlook {
            app,
            namespace
        })
    }

    // Root folder, first name on path, should be the base address in Outlook
    pub fn get_folder(&self, path_to_folder : Vec<&str>) -> Result<Option<Folder>, WinError> {
        let mut folder_names = path_to_folder.into_iter();

        let top_folder_name = folder_names.next().expect("Folder chain cannot be empty");
        let mut subfolder = match Folder(self.namespace.clone()).get_subfolder(top_folder_name)? {
            None => return Ok(None),
            Some(folder) => folder,
        };

        for subfolder_name in folder_names {
            subfolder = match subfolder.get_subfolder(subfolder_name)? {
                None => return Ok(None),
                Some(folder) => folder,
            };
        };  
        Ok(Some(subfolder))
    }
}

impl HasDispatch for Outlook {
    fn dispatch(&self) -> &IDispatch {
        &self.app
    }
}

#[derive(Clone)]
pub struct Folder(pub IDispatch);

impl Folder {
    pub fn get_subfolder(&self, folder_name : &str) -> Result<Option<Folder>, WinError> {
        let params = DISPPARAMS {
            cArgs : 1,
            rgvarg : &mut VARIANT::from(folder_name),
            ..Default::default()
        };

        let dispid = self.get_dispid("Folders")?;
        let mut folder = VARIANT::default();

        unsafe {
            self.0.Invoke(
                dispid,
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_PROPERTYGET,
                &params,
                Some(&mut folder),
                None,
                None,
            ).map_err(|e| WinError::Internal(e))?
        }

        return Ok(Some(Folder(IDispatch::try_from(&folder).expect("couldnt cast VARIANT to folder dispatch"))));
    }

    pub(crate) fn subfolder_names(&self) -> Result<Vec<String>, WinError> {
        let mut foldernames = vec![];
        let subfolders = match self.prop("Folders") {
            Ok(TypedVariant::Dispatch(d)) => d,
            Ok(result) => return Err(WinError::VariantError(VariantError::Mismatch { method: "Folders".to_string(), result })),
            Err(e) => return Err(e),
        };

        let Ok(TypedVariant::Dispatch(first_folder)) = subfolders.call("GetFirst", Invocation::Method, None) else {
            return Ok(foldernames);
        };

        match first_folder.prop("Name")? {
            TypedVariant::Bstr(name) => foldernames.push(name.to_string()),
            result => return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result})),
        };

        loop {
            match subfolders.call("GetNext", Invocation::Method, None) {
                Ok(TypedVariant::Dispatch(subfolder)) => {
                    match subfolder.prop("Name")? {
                        TypedVariant::Bstr(name) => foldernames.push(name.to_string()),
                        result => return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result})),
                    }
                },
                Ok(result) => {
                    return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result}))
                }
                Err(WinError::VariantError(VariantError::NullPointer)) => break, // "Iterator" exhausted
                Err(e) => return Err(e),
            }
        }
        Ok(foldernames)
    }

    pub fn count(&self) -> Option<usize> {
        let TypedVariant::Dispatch(items) = self.prop("Items").expect("Folder should have Items property") else {
            return None
        };

        match items.prop("Count") {
            Ok(TypedVariant::Int32(count)) => Some(count as usize),
            Err(WinError::VariantError(VariantError::NullPointer)) => return Some(0),
            _ => None,
        }
    }

    pub fn emails(&self) -> Result<Vec<MailItem>, WinError> {
        let mut items = Vec::with_capacity(self.count().unwrap_or(0));

        items.extend(self.iter()?);
        Ok(items)
    }

    fn iter(&self) -> Result<MailItemIterator, WinError> {
        match self.prop("Items")? {
            TypedVariant::Dispatch(d) => Ok(MailItemIterator(d,true)),
            result => Err(WinError::VariantError(VariantError::Mismatch { method: "Items".to_string(), result })),
        }
    }
}

impl HasDispatch for Folder {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}

pub struct MailItemIterator(IDispatch,bool);

impl HasDispatch for MailItemIterator {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}

impl Iterator for MailItemIterator {
    type Item = MailItem;

    fn next(&mut self) -> Option<Self::Item> {
        let method = if self.1 {
            self.1 = false;
            "GetFirst"
        } else {
            "GetNext"
        };
        match self.call_evil(method, Invocation::Method, None) {
            Ok(EvilVariant { vt : 9, union: 0, .. })  => return None, // End of iterator is nullpointer
            Ok(evil) if evil.vt == 9 => return Some(MailItem(IDispatch::from(evil))),
            Ok(result) => panic!("Expected MailItem Dispatch while iterating, found {:?}", result),
            Err(e) => panic!("MailItem Iterator failed with: {:?}", e),
        };
    }
}

pub struct MailItem(pub IDispatch);

impl MailItem {
    pub fn move_to(&self, target : &Folder) -> Result<(), WinError> {
        let params = DISPPARAMS {
            rgvarg : &mut VARIANT::from(target.0.clone()),
            cArgs : 1,
            ..Default::default()
        };

        self.call("Move", Invocation::Method, Some(params))?;
        
        Ok(())
    }

    // String properties
    fn string_property(&self, name : &str) -> Result<String, WinError> {
        match self.prop(name)? {
            TypedVariant::Bstr(string) => Ok(string.to_string()),
            result => Err(WinError::VariantError(VariantError::Mismatch { method: name.to_string(), result })),
        }
    }

    pub fn subject(&self) -> Result<String, WinError> { // String
        self.string_property("Subject")
    }

    pub fn body(&self) -> Result<String, WinError> { // String
        self.string_property("Body")
    }

    fn received_time(&self) { // ????

    }

    fn sender_address(&self) { // String

    }
}

impl HasDispatch for MailItem {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}