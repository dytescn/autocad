use dyteslogs::logs::LogError;
use windows::Win32::System::Threading::Sleep;
use windows::Win32::System::Variant;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::layouts::IAcadLayout;
use crate::sdk::plot;
use super::plotcfg;
use std::thread::sleep;
use std::time::Duration;
use super::typing::PrintPont;

// 导出预览图
pub fn batch_print_file(src_path:&str,cur_version:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None,).log_error("init com error").unwrap();
    }
        let app = Application::new(&cur_version).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actLayout = active_doc.get_ActiveLayout().unwrap();
        let res = plotcfg::print_plot_default(&actLayout);
        let actplot = active_doc.get_Plot().unwrap();
        let allblock = actLayout.get_Block().unwrap();
        let count = allblock.get_Count();
        let mut allprintpoint:Vec<PrintPont> = Vec::new();
        let mut sum = 1;
        println!("{:?}",count);
        for i in 0..count  {
            let act_enty = allblock.Item(i).unwrap();
            let enty_type = act_enty.get_EntityType();           
            if enty_type==7 {
                let is_arr = act_enty.get_HasAttributes();
                if !is_arr {
                    continue;
                }
                let box_arr = act_enty.get_GetBoundingBox();
                let data1_arr = box_arr[0].to_f64_array().unwrap();
                let data1_arr_x = data1_arr[0] as i64 as f64;
                let data1_arr_y = data1_arr[1] as i64 as f64;
                let data2_arr = box_arr[1].to_f64_array().unwrap();
                let data2_arr_x = data2_arr[0] as i64 as f64;
                let data2_arr_y = data2_arr[1] as i64 as f64;
                let res = actLayout.SetWindowToPlot([data2_arr_x-3.00,data2_arr_y-3.00].to_vec(),[data1_arr_x+3.00,data1_arr_y+3.00].to_vec());
                if res.is_err() {
                    return;
                }
                let file_src = format!("{}_{}.jpg",src_path,sum);
                actplot.PlotToFile(&file_src,""); 
                sum= sum +1;
                // println!("{:?}",box_arr.len());
            }
        }
    unsafe {
        Com::CoUninitialize();
    }
    return;
}