use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadDatabasePreferences{
    disp:ComObject
}

impl IAcadDatabasePreferences{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn get_SolidFill(){
        
    }
    fn put_SolidFill(){
        
    }
    fn get_XRefEdit(){
        
    }
    fn put_XRefEdit(){
        
    }
    fn get_XRefLayerVisibility(){
        
    }
    fn put_XRefLayerVisibility(){
        
    }
    fn get_OLELaunch(){
        
    }
    fn put_OLELaunch(){
        
    }
    fn get_AllowLongSymbolNames(){
        
    }
    fn put_AllowLongSymbolNames(){
        
    }
    fn get_ObjectSortBySelection(){
        
    }
    fn put_ObjectSortBySelection(){
        
    }
    fn get_ObjectSortBySnap(){
        
    }
    fn put_ObjectSortBySnap(){
        
    }
    fn get_ObjectSortByRedraws(){
        
    }
    fn put_ObjectSortByRedraws(){
        
    }
    fn get_ObjectSortByRegens(){
        
    }
    fn put_ObjectSortByRegens(){
        
    }
    fn get_ObjectSortByPlotting(){
        
    }
    fn put_ObjectSortByPlotting(){
        
    }
    fn get_ObjectSortByPSOutput(){
        
    }
    fn put_ObjectSortByPSOutput(){
        
    }
    fn put_ContourLinesPerSurface(){
        
    }
    fn get_ContourLinesPerSurface(){
        
    }
    fn put_DisplaySilhouette(){
        
    }
    fn get_DisplaySilhouette(){
        
    }
    fn put_MaxActiveViewports(){
        
    }
    fn get_MaxActiveViewports(){
        
    }
    fn put_RenderSmoothness(){
        
    }
    fn get_RenderSmoothness(){
        
    }
    fn put_SegmentPerPolyline(){
        
    }
    fn get_SegmentPerPolyline(){
        
    }
    fn put_TextFrameDisplay(){
        
    }
    fn get_TextFrameDisplay(){
        
    }
    fn put_Lineweight(){
        
    }
    fn get_Lineweight(){
        
    }
    fn put_LineWeightDisplay(){
        
    }
    fn get_LineWeightDisplay(){

    }
}