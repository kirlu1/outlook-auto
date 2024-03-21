use windows::{
    core::{IUnknown, Interface, GUID, VARIANT}, 
    Win32::System::Com::{CoCreateInstance, IDispatch}
};

use anyhow::{bail, Result};

use core::cell::OnceCell;
use std::sync::{Arc, Mutex, OnceLock};

use crate::{bstr, co_initialize, common::{dispatch::{HasDispatch, Invocation}, variant::{TypedVariant, VariantError}}, WinError, OBJECT_CONTEXT};


pub struct Outlook(pub IDispatch);

impl Outlook {
    pub fn new() -> Result<Self, WinError> {
        co_initialize();
    
        let class_id = GUID::from("0006F03A-0000-0000-C000-000000000046");
        let raw_ptr = &class_id as *const GUID;
    
        let unknown : IUnknown  =  unsafe { CoCreateInstance(raw_ptr, None, OBJECT_CONTEXT) }.map_err(
          |e| WinError::Internal(e)  
        )?;
        
        let dispatch : IDispatch = unknown.cast().map_err(
            |e| WinError::Internal(e)
        )?;
        Ok(Outlook(dispatch))
    }

    pub(crate) fn get_namespace(&self) -> IDispatch {
        let TypedVariant::Dispatch(namespace) = self.prop("Session")
            .expect("Failed to get namespace property of Application") 
        else {
            panic!("Namespace wrong VARIANT????");
        };
        namespace
    }

    // Root folder, first name on path, should be the base address in Outlook
    pub fn get_folder(&self, path_to_folder : Vec<&str>) -> Result<Option<Folder>, WinError> {
        let namespace = self.get_namespace();

        let mut folder_names = path_to_folder.into_iter();

        let top_folder_name = folder_names.next().expect("Folder chain cannot be empty");
        let mut subfolder = match Folder(namespace).get_subfolder(top_folder_name)? {
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
        &self.0
    }
}

pub struct Folder(pub IDispatch);

impl Folder {
    pub fn get_subfolder(&self, folder_name : &str) -> Result<Option<Folder>, WinError> {
        let subfolders = match self.prop("Folders") {
            Ok(TypedVariant::Dispatch(d)) => d,
            Ok(result) => return Err(WinError::VariantError(VariantError::Mismatch { method: "Folders".to_string(), result })),
            Err(e) => return Err(e),
        };

        let first_folder = match subfolders.call("GetFirst", Invocation::Method, vec![], false)? {
            TypedVariant::Dispatch(dispatch) => dispatch,
            result => return Err(WinError::VariantError(VariantError::Mismatch { method: "GetFirst".to_string(), result })),
        };

        match first_folder.prop("Name")? {
            TypedVariant::Bstr(name) if &name.to_string() == folder_name => return Ok(Some(Folder(first_folder))),
            TypedVariant::Bstr(name) => (),
            result => return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result})),
        };

        loop {
            match subfolders.call("GetNext", Invocation::Method, vec![], false) {
                Ok(TypedVariant::Dispatch(subfolder)) => {
                    match subfolder.prop("Name")? {
                        TypedVariant::Bstr(name) if &name.to_string() == folder_name => return Ok(Some(Folder(subfolder))),
                        TypedVariant::Bstr(name) => (),
                        result => return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result})),
                    };
                },
                Ok(result) => {
                    return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result}))
                }
                Err(WinError::VariantError(VariantError::NullPointer)) => return Ok(None), // "Iterator" exhausted
                Err(e) => return Err(e),
            }
        }
    }

    pub(crate) fn subfolder_names(&self) -> Result<Vec<String>, WinError> {
        let mut foldernames = vec![];
        let subfolders = match self.prop("Folders") {
            Ok(TypedVariant::Dispatch(d)) => d,
            Ok(result) => return Err(WinError::VariantError(VariantError::Mismatch { method: "Folders".to_string(), result })),
            Err(e) => return Err(e),
        };

        let Ok(TypedVariant::Dispatch(first_folder)) = subfolders.call("GetFirst", Invocation::Method, vec![], false) else {
            return Ok(foldernames);
        };

        match first_folder.prop("Name")? {
            TypedVariant::Bstr(name) => foldernames.push(name.to_string()),
            result => return Err(WinError::VariantError(VariantError::Mismatch {method : "Name".to_string(), result})),
        };

        loop {
            match subfolders.call("GetNext", Invocation::Method, vec![], false) {
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
}

impl HasDispatch for Folder {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}

pub struct MailItem(IDispatch);

impl MailItem {
    fn move_to(&self, target : &Folder) { // Result<_>

    }

    // String properties
    fn subject(&self) { // String
        let bobo = self.prop("Subject").expect("");
    }

    fn body(&self) { // String

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