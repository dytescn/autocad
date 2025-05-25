use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;


// 保存封面图
pub fn autocad_export_cover(cover_src:String,ver:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
        let app = Application::new(&ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actlayout = active_doc.get_ActiveLayout().unwrap();
        // 设置打印
        let _ =  actlayout.put_ConfigName("PublishToWeb JPG.pc3"); // 选择打印机
        let _ = actlayout.put_StyleSheet("monochrome.ctb"); //选择打印样式
        let _ = actlayout.put_PlotWithLineweights(false);   // 不打印等宽线
        let _ = actlayout.put_CanonicalMediaName("UserDefinedRaster (789.00 x 610.00Pixels)");
        let _ = actlayout.put_PlotRotation(0); // 横向打印
        let _ = actlayout.put_CenterPlot(true); // 居中打印
        let _= actlayout.put_UseStandardScale(true);    // 布满图纸
        let _= actlayout.put_PlotWithPlotStyles(true);  // 依照样式打印
        let _ = actlayout.put_PlotHidden(false);        // 赢藏图纸打印空间
        let _ = actlayout.put_PlotType(0);              // 打印类型
        let actplot = active_doc.get_Plot().unwrap();
        let res = actplot.PlotToFile(&cover_src,"");
        return res;
        let _ = Com::CoUninitialize();
    }
}