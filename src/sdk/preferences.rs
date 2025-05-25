use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::preferencesoutput::IAcadPreferencesOutput;

pub struct IAcadPreferences {
    disp:ComObject
}

impl IAcadPreferences{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn get_Files(){
        
    }
    fn get_Display(){
        
    }
    fn get_OpenSave(){
        
    }
    pub fn get_Output(&self)->Option<IAcadPreferencesOutput>{
        let doc_res = self.disp.get_property("output");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let acad_doc = IAcadPreferencesOutput::new(disp.clone());
                        return Some(acad_doc);
                    }
                    Err(e)=>{
                        return None
                    }
                }
            }
            Err(err)=>{
                return None
            }
        }
    }
    fn get_System(){
        
    }
    fn get_User(){
        
    }
    fn get_Drafting(){
        
    }
    fn get_Selection(){
        
    }
    fn get_Profiles(){

    }
}