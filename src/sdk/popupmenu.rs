use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPopupMenu {
    disp:ComObject
}

impl IAcadPopupMenu{
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
    fn get_NameNoMnemonic(){
        
    }
    fn get_ShortcutMenu(){
        
    }
    fn get_OnMenuBar(){
        
    }
    fn AddMenuItem(){
        
    }
    fn AddSubMenu(){
        
    }
    fn AddSeparator(){
        
    }
    fn InsertInMenuBar(){
        
    }
    fn RemoveFromMenuBar(){
        
    }
    fn get_TagString(){

    }
}