use windows::{
    core::{IUnknown, Interface, GUID}, 
    Win32::System::Com::{CoCreateInstance, IDispatch}
};

use anyhow::Result;

use crate::{co_initialize, common::dispatch::HasDispatch, OBJECT_CONTEXT};


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
}

impl HasDispatch for Outlook {
    fn dispatch(&self) -> &IDispatch {
        &self.0
    }
}