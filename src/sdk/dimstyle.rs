use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadDimStyles{
    disp:ComObject
}

impl IAcadDimStyles{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn Item(){

    }
    fn get_Count(){

    }
    fn get__NewEnum(){

    }
    fn Add(){

    }
}


pub struct IAcadDimStyle {
    disp:ComObject
}

impl IAcadDimStyle{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Name(){

    }
    fn put_Name(){

    }
    fn CopyFrom(){

    }
}