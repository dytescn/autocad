use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;
use crate::ffi::application::do_open;

// 保存封面图
pub fn autocad_export_file(file_src:String,ver:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
        let app = Application::new(ver).unwrap();
        let docs =  app.get_documents();
        let doc =app.get_activedocument().unwrap();
        let doc_file_path = doc.get_FullName();
        if doc_file_path=="" {
           let res = doc.do_saveas(&file_src);
           if res.is_some() {
               return Some(());
           }
        }
        // 存储文件
        let res = doc.do_save();
        if res.is_none() {
            return None;
        }
        // 转移文件
        let res = std::fs::copy(doc_file_path, file_src.clone());
        if res.is_err() {
            println!("{:?}",res.err());
            return None;
        }
        // 打开文件
        let res =  docs.Open(&file_src);
        if res.is_none() {
            return None;
        }
        let res = doc.do_close(&file_src);
        if res.is_none() {
            return None;
        }
        let _ = Com::CoUninitialize();
        return Some(());
    }
}


// 保存封面图
pub fn autocad_export_dxf(file_src:String,ver:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
        let app = Application::new(ver).unwrap();
        let docs =  app.get_documents();
        let doc =app.get_activedocument().unwrap();
        let doc_file_path = doc.get_FullName();
        println!("export dxf open file{:?}",doc_file_path);
        // 存储文件
        let res = doc.do_saveas_dxf(&file_src);
        if res.is_none() {
            return None;
        }
        let res = doc.do_close(&file_src);
        if res.is_none() {
        }
        let _ = Com::CoUninitialize();
        return Some(());
    }
}

// 保存封面图
pub fn autocad_saveas_dwg(file_src:String,ver:&str)->Option<()>{
    unsafe {
        let _ = Com::CoInitialize(None).expect("init com error");
        let app = Application::new(ver).unwrap();
        let docs =  app.get_documents();
        let doc =app.get_activedocument().unwrap();
        // 存储文件
        let doc_file_path = doc.get_FullName();
        // let res = doc.do_saveas(&file_src);
        // 转移文件
        let res = std::fs::copy(doc_file_path, file_src.clone());
        if res.is_err() {
            println!("{:?}",res.err());
            return None;
        }
        let _ = Com::CoUninitialize();
        return Some(());
    }
}


