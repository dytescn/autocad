use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::ffi::layout;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 导出预览图
pub fn layout_plot_preview(src_path:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app = Application::new(&app_ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actlayout = active_doc.get_ActiveLayout().unwrap();
        // 设置打印机
        // let name = actlayout.get_CanonicalMediaName();
        // println!("{:?}",name);
        let _ = actlayout.GetWindowToPlot();
        // let _ =  actlayout.put_ConfigName("PublishToWeb JPG.pc3"); // 选择打印机
        // let _ = actlayout.put_StyleSheet("monochrome.ctb"); //选择打印样式
        // let _ = actlayout.put_PlotWithLineweights(false);   // 不打印等宽线
        // let _ = actlayout.put_CanonicalMediaName("funxA3");
        // let _ = actlayout.put_PlotRotation(0); // 横向打印
        // let _ = actlayout.put_CenterPlot(true); // 居中打印
        // let _= actlayout.put_UseStandardScale(true);    // 布满图纸
        // let _= actlayout.put_PlotWithPlotStyles(true);  // 依照样式打印
        // let _ = actlayout.put_PlotHidden(false);        // 赢藏图纸打印空间
        // let _ = actlayout.put_PlotType(4);              // 打印类型
        // let data1 = vec![119944.0,-34.0];
        // let data2 = vec![149834.0,21426.0];
        // let _ = actlayout.GetWindowToPlot();
        // let _ = actlayout.SetWindowToPlot(data1,data2);
        // let actplot = active_doc.get_Plot().unwrap();
        // actplot.PlotToFile(src_path,"");
        Com::CoUninitialize();
    }
}