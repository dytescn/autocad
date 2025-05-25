use dyteslogs::logs::LogError;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

pub fn  export_comps_cover(cover_src: &str,ver:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
    }
    let app = Application::new(&ver).unwrap();
    let active_doc = app.get_activedocument().unwrap();
    let actlayout = active_doc.get_ActiveLayout().unwrap();
    let actSelect = active_doc.get_ActiveSelectionSet().unwrap();
    let count = actSelect.get_Count();
    if count <1 {
        return None;
    }
    println!("count:{}",count);
    let actEntity = actSelect.Item(0).unwrap();
    let name = actEntity.get_name();
    println!("actEntity:{}",name);
    let _ =  actlayout.put_ConfigName("PublishToWeb JPG.pc3"); // 选择打印机
    let _ = actlayout.put_StyleSheet("monochrome.ctb"); //选择打印样式
    let _ = actlayout.put_PlotWithLineweights(false);   // 不打印等宽线
    let _ = actlayout.put_CanonicalMediaName("UserDefinedRaster (789.00 x 610.00Pixels)");
    let _ = actlayout.put_PlotRotation(0); // 横向打印
    let _ = actlayout.put_CenterPlot(true); // 居中打印
    let _= actlayout.put_UseStandardScale(true);    // 布满图纸
    let _= actlayout.put_PlotWithPlotStyles(true);  // 依照样式打印
    let _ = actlayout.put_PlotHidden(false);        // 赢藏图纸打印空间
    let box_arr = actEntity.get_GetBoundingBox();
    let data1_arr = box_arr[0].to_f64_array().unwrap();
    let data1_arr_x = data1_arr[0] as i64 as f64;
    let data1_arr_y = data1_arr[1] as i64 as f64;
    let data2_arr = box_arr[1].to_f64_array().unwrap();
    let data2_arr_x = data2_arr[0] as i64 as f64;
    let data2_arr_y = data2_arr[1] as i64 as f64;
    let res = actlayout.SetWindowToPlot([data2_arr_x-3.00,data2_arr_y-3.00].to_vec(),[data1_arr_x+3.00,data1_arr_y+3.00].to_vec());
    if res.is_err() {
        return None;
    }
    let actplot = active_doc.get_Plot().unwrap();
    actplot.PlotToFile(&cover_src,""); 
    // 设置打印
    unsafe {
        let _ = Com::CoUninitialize();
    }
    return Some(());
}