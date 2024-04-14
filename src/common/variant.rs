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
    UnsupportedVariant
}

impl Display for VariantError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            &VariantError::Opaque => write!(f, "Internal windows API error"),
            VariantError::NullPointer => write!(f, "Null-pointer in non-empty VARIANT"),
            &VariantError::UnsupportedVariant => write!(f, "VARIANT not represented by enum"),
            VariantError::Mismatch { method, result }  => write!(f, "{} returned {:?}", method, result,),
        }
    }
}

impl Error for VariantError {

}

#[derive(Debug, Default)]
#[repr(C)]
pub(crate) struct EvilVariant {
    pub(crate) vt : u16,
    trash1 : u16,
    trash2 : u16,
    trash3 : u16,
    pub(crate) union : usize,
    rec : usize,
}

impl EvilVariant {
    fn new(vt : u16, union_variant : usize) -> Self {
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

impl From<EvilVariant> for IDispatch {
    fn from(value: EvilVariant) -> Self {
        assert_ne!(value.union, 0);
        assert_eq!(value.vt, 9);

        IDispatch::try_from(&VARIANT::from(value)).expect("EvilVariant not valid IDispatch")
    }
}

impl Drop for EvilVariant {
    fn drop(&mut self) {
        
    }
}

#[repr(u16)]
#[derive(Debug)]
pub(crate) enum TypedVariant {
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

    fn try_from(value: VARIANT) -> Result<Self, Self::Error> {
        let value = if let Ok(dispatch) = IDispatch::try_from(&value) {
            TypedVariant::Dispatch(dispatch)
        } else if let Ok(bstr) = BSTR::try_from(&value) {
            TypedVariant::Bstr(bstr)
        } else if let Ok(num) = i32::try_from(&value) {
            TypedVariant::Int32(num)
        } else {
            return Err(WinError::VariantError(VariantError::UnsupportedVariant))
        };
        Ok(value)
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

fn opt_out_arg() -> VARIANT {
    let ev = EvilVariant {
        vt : 10,
        union : 0x80020004,
        ..Default::default()
    };

    VARIANT::from(ev)
}