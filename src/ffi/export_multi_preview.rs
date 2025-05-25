
use dyteslogs::logs::LogError;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use crate::ffi::plotcfg;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;
use crate::ffi::typing::PrintPont;

// 保存封面图
pub fn autocad_export_preview_point(ver:&str)->Option<Vec<PrintPont>>{
    unsafe {
        // 初始化系统
        let conn_res = Com::CoInitialize(None,);
        if conn_res.is_err() {
            return None;
        }
    }
    let app = Application::new(&ver).unwrap();
    let active_doc_res = app.get_activedocument();
    if active_doc_res.is_none() {
        return None;
    }
    let active_doc = active_doc_res.unwrap();
    let actLayout_res = active_doc.get_ActiveLayout();
    if actLayout_res.is_none() {
        return None;
    }
    let actLayout = actLayout_res.unwrap();
    let allblock = actLayout.get_Block().unwrap();
    let count = allblock.get_Count();
    let mut allprintpoint:Vec<PrintPont> = Vec::new();
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
            let ppoint = PrintPont{
                start_x:data2_arr_x-3.00,
                start_y:data2_arr_y-3.00,
                end_x:data1_arr_x+3.00,
                end_y:data1_arr_y+3.00
            };
            allprintpoint.push(ppoint);
        }
    }
    unsafe {
        Com::CoUninitialize();
    }
    return Some(allprintpoint);
}

