
use dyteslogs::logs::LogError;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use crate::ffi::plotcfg;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;
use crate::ffi::typing::PrintPont;
use std::thread::sleep;
use std::time::Duration;

// 保存封面图
pub fn autocad_export_preview_point(src_path:&str,print_num:usize,ver:&str,ppiont:PrintPont)->Option<()>{
    unsafe {
        // 初始化系统
        let conn_res = Com::CoInitialize(None,);
        if conn_res.is_err() {
            return None;
        }
    }
    let app = Application::new(&ver).unwrap();
    let active_doc = app.get_activedocument().unwrap();
    let actLayout = active_doc.get_ActiveLayout().unwrap();
    let res = plotcfg::print_plot_default(&actLayout);
    let actplot = active_doc.get_Plot().unwrap();
    let data1 = vec![ppiont.start_x,ppiont.start_y];
    let data2 = vec![ppiont.end_x,ppiont.end_y];
    let res = actLayout.SetWindowToPlot(data1.clone(),data2.clone());
    let _= actLayout.put_UseStandardScale(true);    // 布满图纸
    if res.is_err() {
        // sleep(Duration::from_secs(3));  // Sleep for 10 seconds
        println!("set SetWindowToPlot error....");
    }
    let mut do_plot = 1;
    while do_plot==1 {
        let src = format!("{}{}_preview.jpg",src_path,print_num);
        let res = actplot.PlotToFile(&src,"");   
        if res.is_some() {
            // println!("error:{:?}",res.err());
            println!("...");
            do_plot = 2
        }
        sleep(Duration::from_secs(3));  // Sleep for 10 seconds
    }
    unsafe {
        Com::CoUninitialize();
    } 
    return Some(());
}

