use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadMenuGroup {
    disp:ComObject
}

impl IAcadMenuGroup{
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
    fn get_Type(){
        
    }
    fn get_MenuFileName(){
        
    }
    fn get_Menus(){
        
    }
    fn get_Toolbars(){
        
    }
    fn Unload(){
        
    }
    fn Save(){
        
    }
    fn SaveAs(){

    }
}