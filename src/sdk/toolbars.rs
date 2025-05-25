use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadToolbars {
    disp:ComObject
}

impl IAcadToolbars{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn Item(){
        
    }
    fn get__NewEnum(){
        
    }
    fn get_Count(){
        
    }
    fn get_Application(){
        
    }
    fn get_Parent(){
        
    }
    fn get_LargeButtons(){
        
    }
    fn put_LargeButtons(){
        
    }
    fn Add(){

    }
}