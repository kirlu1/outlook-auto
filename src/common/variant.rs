use std::{error::Error, fmt::Display};

use windows::{core::{IUnknown, BSTR, VARIANT}, Win32::System::Com::IDispatch};

use crate::WinError;

#[derive(Debug)]
pub enum VariantError {
    Opaque,
    NullPointer,
    Mismatch {
        method : String,
        result : TypedVariant,
    },
}

impl Display for VariantError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            &VariantError::Opaque => write!(f, "Internal windows API error"),
            VariantError::NullPointer => write!(f, "Null-pointer in non-empty VARIANT"),
            VariantError::Mismatch { method, result }  => write!(f, "{} returned {:?}", method, result,),
        }
    }
}

impl Error for VariantError {

}

#[derive(Debug)]
#[repr(C)]
pub struct EvilVariant {
    pub vt : u16,
    trash1 : u16,
    trash2 : u16,
    trash3 : u16,
    pub union : usize,
    rec : usize,
}

impl EvilVariant {
    pub fn new(vt : u16, union_variant : usize) -> Self {
        EvilVariant {
            vt,
            trash1 : 0,
            trash2 : 0,
            trash3 : 0,
            union : union_variant,
            rec : 0
        }
    }

    fn is_null(&self) -> bool {
        self.union==0
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
pub enum TypedVariant {
    Empty = 0x00,
    Int32(i32) = 0x03,
    Bstr(BSTR) = 0x08,
    Dispatch(IDispatch) = 0x09,
    Unknown(IUnknown) = 0x0D
}

impl TypedVariant {
    fn as_u16(&self) -> u16 {
        match self {
            TypedVariant::Empty => 0x00,
            TypedVariant::Int32(_) => 0x03,
            TypedVariant::Bstr(_) => 0x08,
            TypedVariant::Dispatch(_) => 0x09,
            TypedVariant::Unknown(_) => 0x0D,
        }
    }
}

impl TryFrom<VARIANT> for TypedVariant {
    type Error = WinError;

    fn try_from(value: VARIANT) -> Result<TypedVariant, WinError> {
        let evil_variant = EvilVariant::from(value);
        match evil_variant.vt {
            // union variant may be pointer *OR* value
            0x00 => Ok(TypedVariant::Empty),
            _ if evil_variant.is_null() => Err(WinError::VariantError(VariantError::NullPointer)),
            0x03 => Ok(TypedVariant::Int32(evil_variant.union as i32)),
            0x08 => Ok(TypedVariant::Bstr(unsafe { std::mem::transmute::<usize, BSTR>(evil_variant.union) }.clone())),
            0x09 => Ok(TypedVariant::Dispatch(unsafe { std::mem::transmute::<&usize, &IDispatch>(&evil_variant.union) }.clone())),
            0x0D => Ok(TypedVariant::Unknown(unsafe { std::mem::transmute::<&usize, &IUnknown>(&evil_variant.union) }.clone())),
            x => panic!("Strange VType: {}", x)
        }
    }
}

impl From<TypedVariant> for EvilVariant {
    fn from(value: TypedVariant) -> EvilVariant {
        let vt : u16 = value.as_u16();

        // This might be very illegal
        let union_variant = match value {
            TypedVariant::Empty => 0,
            TypedVariant::Int32(num) => num as usize,
            TypedVariant::Bstr(bstr) => unsafe { std::mem::transmute::<BSTR, usize>(bstr) },
            TypedVariant::Dispatch(dispatch) => unsafe { std::mem::transmute::<&IDispatch, usize>(&dispatch) },
            TypedVariant::Unknown(unknown) => unsafe { std::mem::transmute::<&IUnknown, usize>(&unknown) },
        };
        EvilVariant::new(vt, union_variant)
    }
}


impl From<TypedVariant> for VARIANT {
    fn from(value: TypedVariant) -> VARIANT {
        let x = EvilVariant::from(value);
        let result = x.into();
        result
    }
}


pub fn opt_out_arg() -> VARIANT {
    let ev = EvilVariant {
        vt : 10,
        union : 0x80020004,
        ..Default::default()
    };

    VARIANT::from(ev)
}