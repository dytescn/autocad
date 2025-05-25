use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadSummaryInfo{
    disp:ComObject
}

impl IAcadSummaryInfo{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Author(){
        
    }
    fn put_Author(){
        
    }
    fn get_Comments(){
        
    }
    fn put_Comments(){
        
    }
    fn get_HyperlinkBase(){
        
    }
    fn put_HyperlinkBase(){
        
    }
    fn get_Keywords(){
        
    }
    fn put_Keywords(){
        
    }
    fn get_LastSavedBy(){
        
    }
    fn put_LastSavedBy(){
        
    }
    fn get_RevisionNumber(){
        
    }
    fn put_RevisionNumber(){
        
    }
    fn get_Subject(){
        
    }
    fn put_Subject(){
        
    }
    fn get_Title(){
        
    }
    fn put_Title(){
        
    }
    fn NumCustomInfo(){
        
    }
    fn GetCustomByIndex(){
        
    }
    fn GetCustomByKey(){
        
    }
    fn SetCustomByIndex(){
        
    }
    fn SetCustomByKey(){
        
    }
    fn AddCustomInfo(){
        
    }
    fn RemoveCustomByIndex(){
        
    }
    fn RemoveCustomByKey(){

    }
}