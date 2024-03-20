use std::ptr;

use windows::{core::{IUnknown, BSTR, VARIANT}, Win32::System::Com::IDispatch};

#[derive(Debug)]
#[repr(C)]
pub struct EvilVariant {
    pub vt : u16,
    trash1 : u16,
    trash2 : u16,
    trash3 : u16,
    pub union : u64,
    rec : usize,
}

impl EvilVariant {
    pub fn new(vt : u16, union_variant : u64) -> Self {
        EvilVariant {
            vt,
            trash1 : 0,
            trash2 : 0,
            trash3 : 0,
            union : union_variant,
            rec : 0
        }
    }
}

impl From<VARIANT> for EvilVariant {
    fn from(value: VARIANT) -> Self {
        unsafe {
            std::mem::transmute(value)
        }
    }
}

impl From<EvilVariant> for VARIANT {
    fn from(value: EvilVariant) -> Self {
        unsafe {
            std::mem::transmute(value)
        }
    }
}

impl Drop for EvilVariant {
    fn drop(&mut self) {
        
    }
}

#[repr(u16)]
pub enum SafeVariant {
    Bstr(BSTR) = 0x08,
    Dispatch(IDispatch) = 0x09,
    Unknown(IUnknown) = 0x0D
}

impl From<VARIANT> for SafeVariant {
    fn from(value: VARIANT) -> Self {
        let evil_variant = EvilVariant::from(value);
        match evil_variant.vt {
            // Some union variants are pointers, some are values 
            0x08 => SafeVariant::Bstr(unsafe { std::mem::transmute::<u64, BSTR>(evil_variant.union) }.clone()),
            0x09 => SafeVariant::Dispatch(unsafe { std::mem::transmute::<&u64, &IDispatch>(&evil_variant.union) }.clone()),
            0x0D => SafeVariant::Unknown(unsafe { std::mem::transmute::<&u64, &IUnknown>(&evil_variant.union) }.clone()),
            x => panic!("Strange VT: {}", x)
        }
    }
}

// let dispatch_ref : &IDispatch = unsafe { std::mem::transmute(&public.union) };
// Ok(DispatchObject(dispatch_ref.clone()))