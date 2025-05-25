use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesOpenSave {
    disp:ComObject
}

impl IAcadPreferencesOpenSave{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_SavePreviewThumbnail(){
        
    }
    fn get_SavePreviewThumbnail(){
        
    }
    fn put_IncrementalSavePercent(){
        
    }
    fn get_IncrementalSavePercent(){
        
    }
    fn put_AutoSaveInterval(){
        
    }
    fn get_AutoSaveInterval(){
        
    }
    fn put_CreateBackup(){
        
    }
    fn get_CreateBackup(){
        
    }
    fn put_FullCRCValidation(){
        
    }
    fn get_FullCRCValidation(){
        
    }
    fn put_LogFileOn(){
        
    }
    fn get_LogFileOn(){
        
    }
    fn put_TempFileExtension(){
        
    }
    fn get_TempFileExtension(){
        
    }
    fn put_XrefDemandLoad(){
        
    }
    fn get_XrefDemandLoad(){
        
    }
    fn put_DemandLoadARXApp(){
        
    }
    fn get_DemandLoadARXApp(){
        
    }
    fn put_ProxyImage(){
        
    }
    fn get_ProxyImage(){
        
    }
    fn put_ShowProxyDialogBox(){
        
    }
    fn get_ShowProxyDialogBox(){
        
    }
    fn put_AutoAudit(){
        
    }
    fn get_AutoAudit(){
        
    }
    fn put_SaveAsType(){
        
    }
    fn get_SaveAsType(){
        
    }
    fn get_MRUNumber(){

    }
}