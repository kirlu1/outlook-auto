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
#[derive(Debug)]
pub enum SafeVariant {
    Int32(i32) = 0x03,
    Bstr(BSTR) = 0x08,
    Dispatch(IDispatch) = 0x09,
    Unknown(IUnknown) = 0x0D
}

impl SafeVariant {
    fn as_u16(&self) -> u16 {
        match self {
            SafeVariant::Int32(_) => 0x03,
            SafeVariant::Bstr(_) => 0x08,
            SafeVariant::Dispatch(_) => 0x09,
            SafeVariant::Unknown(_) => 0x0D,
        }
    }
}

impl From<VARIANT> for SafeVariant {
    fn from(value: VARIANT) -> Self {
        let evil_variant = EvilVariant::from(value);
        match evil_variant.vt {
            // union variant may be pointer *OR* value
            0x03 => SafeVariant::Int32(evil_variant.union as i32),
            0x08 => SafeVariant::Bstr(unsafe { std::mem::transmute::<u64, BSTR>(evil_variant.union) }.clone()),
            0x09 => SafeVariant::Dispatch(unsafe { std::mem::transmute::<&u64, &IDispatch>(&evil_variant.union) }.clone()),
            0x0D => SafeVariant::Unknown(unsafe { std::mem::transmute::<&u64, &IUnknown>(&evil_variant.union) }.clone()),
            x => panic!("Strange VType: {}", x)
        }
    }
}


impl From<SafeVariant> for VARIANT {
    fn from(value: SafeVariant) -> VARIANT {
        let vt : u16 = value.as_u16();

        // This might be very illegal
        let union_variant = match value {
            SafeVariant::Int32(num) => num as u64,
            SafeVariant::Bstr(bstr) => unsafe { std::mem::transmute::<&BSTR, u64>(&bstr) },
            SafeVariant::Dispatch(dispatch) => unsafe { std::mem::transmute::<&IDispatch, u64>(&dispatch) },
            SafeVariant::Unknown(unknown) => unsafe { std::mem::transmute::<&IUnknown, u64>(&unknown) },
        };

        let evil_variant = EvilVariant::new(vt, union_variant);

        evil_variant.into()         
    }
}