use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadToolbarItem {
    disp:ComObject
}

impl IAcadToolbarItem{
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
    fn get_Name(){

    }
    fn put_Name(){

    }
    fn get_TagString(){

    }
    fn put_TagString(){

    }
    fn get_Type(){

    }
    fn get_Flyout(){

    }
    fn get_Macro(){

    }
    fn put_Macro(){

    }
    fn get_Index(){

    }
    fn get_HelpString(){

    }
    fn put_HelpString(){

    }
    fn GetBitmaps(){

    }
    fn SetBitmaps(){

    }
    fn AttachToolbarToFlyout(){

    }
    fn Delete(){

    }
    fn get_CommandDisplayName(){

    }
    fn put_CommandDisplayName(){

    }
}