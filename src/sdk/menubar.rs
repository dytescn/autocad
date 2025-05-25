use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadMenuBar {
    disp:ComObject
}

impl IAcadMenuBar{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn Item(&self,index:u8){
        
    }
    fn get_Count(){

    }   
    fn get_Application(){

    }
    fn get_Parent(){

    }
}