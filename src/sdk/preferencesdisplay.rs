use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPreferencesDisplay {
    disp:ComObject
}

impl IAcadPreferencesDisplay{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){

    }
    
    fn put_LayoutDisplayMargins(){

    }
    fn get_LayoutDisplayMargins(){

    }
    fn put_LayoutDisplayPaper(){

    }
    fn get_LayoutDisplayPaper(){

    }
    fn put_LayoutDisplayPaperShadow(){

    }
    fn get_LayoutDisplayPaperShadow(){

    }
    fn put_LayoutShowPlotSetup(){

    }
    fn get_LayoutShowPlotSetup(){

    }
    fn put_LayoutCreateViewport(){

    }
    fn get_LayoutCreateViewport(){

    }
    fn put_DisplayScrollBars(){

    }
    fn get_DisplayScrollBars(){

    }
    fn put_DisplayScreenMenu(){

    }
    fn get_DisplayScreenMenu(){

    }
    fn put_CursorSize(){

    }
    fn get_CursorSize(){

    }
    fn put_DockedVisibleLines(){

    }
    fn get_DockedVisibleLines(){

    }
    fn put_ShowRasterImage(){

    }
    fn get_ShowRasterImage(){

    }
    fn put_GraphicsWinModelBackgrndColor(){

    }
    fn get_GraphicsWinModelBackgrndColor(){

    }
    fn put_ModelCrosshairColor(){

    }
    fn get_ModelCrosshairColor(){

    }
    fn put_GraphicsWinLayoutBackgrndColor(){

    }
    fn get_GraphicsWinLayoutBackgrndColor(){

    }
    fn put_TextWinBackgrndColor(){

    }
    fn get_TextWinBackgrndColor(){

    }
    fn put_TextWinTextColor(){

    }
    fn get_TextWinTextColor(){

    }
    fn put_LayoutCrosshairColor(){

    }
    fn get_LayoutCrosshairColor(){

    }
    fn put_AutoTrackingVecColor(){

    }
    fn get_AutoTrackingVecColor(){

    }
    fn put_TextFont(){

    }
    fn get_TextFont(){

    }
    fn put_TextFontStyle(){

    }
    fn get_TextFontStyle(){

    }
    fn put_TextFontSize(){

    }
    fn get_TextFontSize(){

    }
    fn put_HistoryLines(){

    }
    fn get_HistoryLines(){

    }
    fn put_MaxAutoCADWindow(){

    }
    fn get_MaxAutoCADWindow(){

    }
    fn put_DisplayLayoutTabs(){

    }
    fn get_DisplayLayoutTabs(){

    }
    fn put_ImageFrameHighlight(){

    }
    fn get_ImageFrameHighlight(){

    }
    fn put_TrueColorImages(){

    }
    fn get_TrueColorImages(){

    }
    fn put_XRefFadeIntensity(){

    }
    fn get_XRefFadeIntensity(){

    }
}