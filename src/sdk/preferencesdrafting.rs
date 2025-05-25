use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesDrafting {
    disp:ComObject
}

impl IAcadPreferencesDrafting{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_AutoSnapMarker(){
        
    }
    fn get_AutoSnapMarker(){
        
    }
    fn put_AutoSnapMagnet(){

    }
    
    fn get_AutoSnapMagnet(){
        
    }
    fn put_AutoSnapTooltip(){
        
    }
    fn get_AutoSnapTooltip(){
        
    }
    fn put_AutoSnapAperture(){
        
    }
    fn get_AutoSnapAperture(){
        
    }
    fn put_AutoSnapApertureSize(){
        
    }
    fn get_AutoSnapApertureSize(){
        
    }    
    fn put_AutoSnapMarkerColor(){
        
    }
    fn get_AutoSnapMarkerColor(){
        
    }
    fn put_AutoSnapMarkerSize(){
        
    }
    fn get_AutoSnapMarkerSize(){
        
    }
    fn put_PolarTrackingVector(){
        
    }
    fn get_PolarTrackingVector(){
        
    }
    fn put_FullScreenTrackingVector(){
        
    }
    fn get_FullScreenTrackingVector(){
        
    }
    fn put_AutoTrackTooltip(){
        
    }
    fn get_AutoTrackTooltip(){
        
    }
    fn put_AlignmentPointAcquisition(){
        
    }
    fn get_AlignmentPointAcquisition(){

    }
}