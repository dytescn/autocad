use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPopupMenuItem {
    disp:ComObject
}

impl IAcadPopupMenuItem{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){

    }
    fn get_Parent(){

    }
    fn get_Label(){

    }
    fn put_Label(){

    }
    fn get_TagString(){

    }
    fn put_TagString(){

    }
    fn get_Enable(){

    }
    fn put_Enable(){

    }
    fn get_Check(){

    }
    fn put_Check(){

    }
    fn get_Type(){

    }
    fn get_SubMenu(){

    }
    fn get_Macro(){

    }
    fn put_Macro(){

    }
    fn get_Index(){

    }
    fn get_Caption(){

    }
    fn get_HelpString(){

    }
    fn put_HelpString(){

    }
    fn Delete(){

    }
    fn get_EndSubMenuLevel(){

    }
    fn put_EndSubMenuLevel(){

    }
}