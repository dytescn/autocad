use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesUser {
    disp:ComObject
}

impl IAcadPreferencesUser{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_KeyboardAccelerator(){
        
    }
    fn get_KeyboardAccelerator(){
        
    }
    fn put_KeyboardPriority(){
        
    }
    fn get_KeyboardPriority(){
        
    }
    fn put_HyperlinkDisplayCursor(){
        
    }
    fn get_HyperlinkDisplayCursor(){
        
    }
    fn put_ADCInsertUnitsDefaultSource(){
        
    }
    fn get_ADCInsertUnitsDefaultSource(){
        
    }
    fn put_ADCInsertUnitsDefaultTarget(){
        
    }
    fn get_ADCInsertUnitsDefaultTarget(){
        
    }
    fn put_ShortCutMenuDisplay(){
        
    }
    fn get_ShortCutMenuDisplay(){
        
    }
    fn put_SCMDefaultMode(){
        
    }
    fn get_SCMDefaultMode(){
        
    }
    fn put_SCMEditMode(){
        
    }
    fn get_SCMEditMode(){
        
    }
    fn put_SCMCommandMode(){
        
    }
    fn get_SCMCommandMode(){
        
    }
    fn put_SCMTimeMode(){
        
    }
    fn get_SCMTimeMode(){
        
    }
    fn put_SCMTimeValue(){
        
    }
    fn get_SCMTimeValue(){

    }
}