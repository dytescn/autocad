use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 保存封面图
pub fn autocad_import_dwg(file_src:&str,ver:&str)->Option<()>{
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            println!("CoInitialize failed: {:?}", res.err());
            return None;
        }
    }
    let app = Application::new(ver).unwrap();
    let doc =app.get_activedocument().unwrap();
    let database = doc.get_Database().unwrap();
    let modelspace = database.get_modelspace().unwrap();
    let res = modelspace.InsertBlock(&file_src,vec![0.0,0.0,0.0],1,0);
    if res.is_none() {
        println!("failed");
        return None;
    }
    return Some(());
}