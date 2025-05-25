use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::ffi::layout;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 导出预览图
pub fn layer_plot_preview(src_path:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app = Application::new(&app_ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        // let actlayout = active_doc.get_ActiveLayout().unwrap();
        // let block = actlayout.get_Block().unwrap();
        let database = active_doc.get_Database().unwrap();
        let layers = database.get_Layers().unwrap();
        let sum = layers.get_Count();
        let layer = layers.Item(0).unwrap();
        let plotname = layer.get_PlotStyleName();
        println!("sum::{:?},plotname:{:?}",sum,plotname);
    }
}