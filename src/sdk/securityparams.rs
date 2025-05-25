use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;


pub struct IAcadSecurityParams {
    disp:ComObject
}

impl IAcadSecurityParams{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn put_Action(){
        
    }
    fn get_Action(){
        
    }
    fn put_Password(){
        
    }
    fn get_Password(){
        
    }
    fn put_ProviderType(){
        
    }
    fn get_ProviderType(){
        
    }
    fn put_ProviderName(){
        
    }
    fn get_ProviderName(){
        
    }
    fn put_Algorithm(){
        
    }
    fn get_Algorithm(){
        
    }
    fn put_KeyLength(){
        
    }
    fn get_KeyLength(){
        
    }
    fn put_Subject(){
        
    }
    fn get_Subject(){
        
    }
    fn put_Issuer(){
        
    }
    fn get_Issuer(){
        
    }
    fn put_SerialNumber(){
        
    }
    fn get_SerialNumber(){
        
    }
    fn put_Comment(){
        
    }
    fn get_Comment(){
        
    }
    fn put_TimeServer(){
        
    }
    fn get_TimeServer(){

    }
}