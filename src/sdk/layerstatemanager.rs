use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadLayerStateManager {
    disp:ComObject
}

impl IAcadLayerStateManager{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn SetDatabase(){
        
    }
    fn put_Mask(){
        
    }
    fn get_Mask(){
        
    }
    fn Save(){
        
    }
    fn Restore(){
        
    }
    fn Delete(){
        
    }
    fn Rename(){
        
    }
    fn Import(){
        
    }
    fn Export(){

    }
}