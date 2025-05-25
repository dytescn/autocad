use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadToolbar {
    disp:ComObject
}

impl IAcadToolbar{
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
    fn get_Name(){
        
    }
    fn put_Name(){
        
    }
    fn get_Visible(){
        
    }
    fn put_Visible(){
        
    }
    fn get_DockStatus(){
        
    }
    fn get_LargeButtons(){
        
    }
    fn get_left(){
        
    }
    fn put_left(){
        
    }
    fn get_top(){
        
    }
    fn put_top(){
        
    }
    fn get_Width(){
        
    }
    fn get_Height(){
        
    }
    fn get_FloatingRows(){
        
    }
    fn put_FloatingRows(){
        
    }
    fn get_HelpString(){
        
    }
    fn put_HelpString(){
        
    }
    fn AddToolbarButton(){
        
    }
    fn AddSeparator(){
        
    }
    fn Dock(){
        
    }
    fn Float(){
        
    }
    fn Delete(){

    }
    fn get_TagString(){

    }
}