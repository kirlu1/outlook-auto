use windows::{
    core::{IUnknown, Interface, GUID, VARIANT}, 
    Win32::System::Com::{CoCreateInstance, IDispatch}
};

use anyhow::{bail, Result};

use core::cell::OnceCell;
use std::sync::{Arc, Mutex, OnceLock};

use crate::{bstr, co_initialize, common::{dispatch::{HasDispatch, Invocation}, variant::SafeVariant}, OBJECT_CONTEXT};


pub struct Outlook(pub IDispatch);

impl Outlook {
    pub fn new() -> Result<Self> {
        co_initialize();
    
        let class_id = GUID::from("0006F03A-0000-0000-C000-000000000046");
        let raw_ptr = &class_id as *const GUID;
    
        let unknown : IUnknown  = unsafe {
            CoCreateInstance(raw_ptr, None, OBJECT_CONTEXT)
        }?;
        
        let dispatch : IDispatch = unknown.cast()?;
        Ok(Outlook(dispatch))
    }

    fn get_namespace(&self) -> IDispatch {
        let SafeVariant::Dispatch(namespace) = self.prop("Session")
            .expect("Failed to get namespace property of Application") 
        else {
            panic!("Namespace wrong VARIANT????");
        };
        namespace
    }

    pub fn get_folder(&self, path_to_folder : Vec<&str>) -> Result<Folder> {
        let namespace = self.get_namespace();

        let mut folder_names = path_to_folder.into_iter();

        let top_folder_name = folder_names.next().expect("Folder chain cannot be empty");
        let Some(mut subfolder) = Folder(namespace).get_subfolder(top_folder_name) else {
            bail!("Cannot find folder {}", top_folder_name)
        };

        for folder_arg in folder_names {
            subfolder = match subfolder.get_subfolder(folder_arg) {
                Some(x) => x,
                None => bail!("Cannot find folder {}", folder_arg),
            };
        };  
        Ok(subfolder)
    }
}

impl HasDispatch for Outlook {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}

pub struct Folder(IDispatch);

impl Folder {
    fn get_subfolder(&self, folder_name : &str) -> Option<Folder> {
        let bstr = bstr(folder_name).expect("String conversion should not fail");
        let safevar = SafeVariant::Bstr(bstr);
        let arg = VARIANT::from(safevar);
        
        let Ok(SafeVariant::Dispatch(subfolders)) = self.prop("Folders") else {
            return None
        };
        let Ok(SafeVariant::Dispatch(target_folder)) = subfolders.call("Item", Invocation::PropertyGet, vec![arg]) else {
            return None
        };
        
        Some(Folder(target_folder))
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