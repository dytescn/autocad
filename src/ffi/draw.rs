use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 保存封面图
pub fn autocad_draw(ver:&str)->Option<()>{
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            return None;
        }
    }
    let app = Application::new(ver).unwrap();
    let doc =app.get_activedocument().unwrap();
    let database = doc.get_Database().unwrap();
    let modelspace = database.get_modelspace().unwrap();
    // 画圆
    let res = modelspace.AddCircle(vec![0.0,0.0,0.0],10000);
    if res.is_none() {
        return None;
    }
    return Some(());
}