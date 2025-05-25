use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesProfiles {
    disp:ComObject
}

impl IAcadPreferencesProfiles{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_ActiveProfile(){
        
    }
    fn get_ActiveProfile(){
        
    }
    fn ImportProfile(){
        
    }
    fn ExportProfile(){
        
    }
    fn DeleteProfile(){
        
    }
    fn ResetProfile(){
        
    }
    fn RenameProfile(){
        
    }
    fn CopyProfile(){
        
    }
    fn GetAllProfileNames(){

    }
}