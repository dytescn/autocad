use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesSystem {
    disp:ComObject
}

impl IAcadPreferencesSystem{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_SingleDocumentMode(){
        
    }
    fn get_SingleDocumentMode(){
        
    }
    fn put_DisplayOLEScale(){
        
    }
    fn get_DisplayOLEScale(){
        
    }
    fn put_StoreSQLIndex(){
        
    }
    fn get_StoreSQLIndex(){
        
    }
    fn put_TablesReadOnly(){
        
    }
    fn get_TablesReadOnly(){
        
    }
    fn put_EnableStartupDialog(){
        
    }
    fn get_EnableStartupDialog(){
        
    }
    fn put_BeepOnError(){
        
    }
    fn get_BeepOnError(){
        
    }
    fn put_ShowWarningMessages(){
        
    }
    fn get_ShowWarningMessages(){
        
    }
    fn put_LoadAcadLspInAllDocuments(){
        
    }
    fn get_LoadAcadLspInAllDocuments(){

    }
}