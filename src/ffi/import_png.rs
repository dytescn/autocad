use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 保存封面图
pub fn autocad_import_png(file_src:&str,scale:i32,ver:&str)->Option<()>{
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            println!("1");
            return None;
        }
    }
    let app_res = Application::new(ver);
    if app_res.is_none() {
        println!("2");
        return None;
    }
    let app = app_res.unwrap();
    let doc =app.get_activedocument().unwrap();
    let database = doc.get_Database().unwrap();
    let modelspace = database.get_modelspace().unwrap();
    let res = modelspace.AddRaster(&file_src,vec![0.0,0.0,0.0],scale,0);
    if res.is_none() {
        return None;
    }

    return Some(());
}