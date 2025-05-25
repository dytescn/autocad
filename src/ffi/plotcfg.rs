// 打印配置
use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::{application::Application, layouts::IAcadLayout, plot};

// 导出预览图
pub fn plot_cfg(src_path:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app = Application::new(&app_ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actlayout = active_doc.get_ActiveLayout().unwrap();
        println!("{:?}","hello");
        // 打印相关参数
        // 打印类型
        let cfg_name = actlayout.get_ConfigName();
        // 打印媒介尺寸
        let medianame = actlayout.get_CanonicalMediaName();
        // 纸张的像素
        let paper_unit = actlayout.get_PaperUnits();
        println!("cfgname:{:?},medianame:{:?},paperunit:{:?}",cfg_name,medianame,paper_unit); 
        // // 打印预览框
        // let plot_border = actlayout.get_PlotViewportBorders();
        // let showstyle = actlayout.get_ShowPlotStyles();
        // let rotation = actlayout.get_PlotOrigin();
        // println!("plot_border:{:?},showstyle:{:?},Rotation:{:?}",plot_border,showstyle,rotation);

        // let window = actlayout.GetWindowToPlot();
        Com::CoUninitialize();
    }
}

// // 打印配置
pub fn print_plot_default(actlayout:&IAcadLayout)->Option<()>{
    // 设置打印机
    let _ =  actlayout.put_ConfigName("PublishToWeb JPG.pc3"); // 选择打印机
    let _ = actlayout.put_StyleSheet("monochrome.ctb"); //选择打印样式
    let _ = actlayout.put_PlotWithLineweights(false);   // 不打印等宽线
    let _ = actlayout.put_CanonicalMediaName("DCI_4K_(4096.00_x_2160.00_Pixels)");
    let _ = actlayout.put_PlotRotation(0); // 横向打印
    let _ = actlayout.put_CenterPlot(true); // 居中打印
    let _= actlayout.put_PlotWithPlotStyles(true);  // 依照样式打印
    let _ = actlayout.put_PlotHidden(false);        // 赢藏图纸打印空间
    let _ = actlayout.put_PlotType(4);              // 打印类型
    let _= actlayout.put_UseStandardScale(true);    // 布满图纸
    return Some(());
}

pub fn print_choose_window(ver:&str,start_x:f64,start_y:f64,end_x:f64,end_y:f64)->Option<()>{
    unsafe {
        // 初始化系统
        Com::CoInitialize(None,).log_error("init com error").unwrap();
        let app = Application::new(&ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actLayout = active_doc.get_ActiveLayout().unwrap();
        let actplot = active_doc.get_Plot().unwrap();
        let data1 = vec![start_x,start_y];
        let data2 = vec![end_x,end_y];
        let res = actLayout.SetWindowToPlot(data1,data2);
        if res.is_err() {
            return None;
        }
        return Some(());
    
    }
}