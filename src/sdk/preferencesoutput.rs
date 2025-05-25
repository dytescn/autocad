use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesOutput {
    disp:ComObject
}

impl IAcadPreferencesOutput{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){
        
    }
    fn put_DefaultOutputDevice(){
        
    }
    fn get_DefaultOutputDevice(){
        
    }
    fn put_PrinterSpoolAlert(){
        
    }
    fn get_PrinterSpoolAlert(){
        
    }
    fn put_PrinterPaperSizeAlert(){
        
    }
    fn get_PrinterPaperSizeAlert(){
        
    }
    fn put_PlotLegacy(){
        
    }
    fn get_PlotLegacy(){
        
    }
    fn put_OLEQuality(){
        
    }
    fn get_OLEQuality(){
        
    }
    fn put_UseLastPlotSettings(){
        
    }
    fn get_UseLastPlotSettings(){
        
    }
    fn put_PlotPolicy(){
        
    }
    fn get_PlotPolicy(){
        
    }
    fn put_DefaultPlotStyleTable(){
        
    }
    pub fn get_DefaultPlotStyleTable(&self)->String{
        let info = self.disp.get_property("DefaultPlotStyleTable").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()   
    }
    fn put_DefaultPlotStyleForObjects(){
        
    }
    pub fn get_DefaultPlotStyleForObjects(&self)->String{
        let info = self.disp.get_property("DefaultPlotStyleForObjects").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap() 
    }
    fn put_DefaultPlotStyleForLayer(){
        
    }
    pub fn get_DefaultPlotStyleForLayer(&self)->String{
        let info = self.disp.get_property("DefaultPlotStyleForLayer").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn put_ContinuousPlotLog(){
        
    }
    fn get_ContinuousPlotLog(){
        
    }
    fn put_AutomaticPlotLog(){
        
    }
    fn get_AutomaticPlotLog(){
        
    }
    fn put_DefaultPlotToFilePath(){
        
    }
    fn get_DefaultPlotToFilePath(){

    }
}