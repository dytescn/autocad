use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
// 获取coreldraw信息
pub fn document_info(cur_version:String){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app = Application::new(&cur_version).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        // let full_name = active_doc.get_FullName();
        // println!("doc:{:?}",full_name);
        // let name = active_doc.get_Name();
        // println!("doc:{:?}",name);
        // let path = active_doc.get_Path();
        // println!("doc:{:?}",path);
        active_doc.put_name("hello");
        // documents
        Com::CoUninitialize();
    }
}

// 获取当前文档的数量
pub fn get_document_count(cur_version:String)->usize{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app = Application::new(&cur_version).unwrap();
        let docs = app.get_documents();
        let sum = docs.get_Count();
        return sum as usize;
        Com::CoUninitialize();
    }
}

// 创建文档
pub fn document_create(){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app_res = Application::new(&app_ver);
        if app_res.is_none() {
            return;
        }
        let app = app_res.unwrap();
        let docs = app.get_documents();
        let selects = docs.Add();
        // let sum  =selects.get_Count();
        println!("print task sum:");
        // let select =  selects.Item(0).unwrap();
        // let hr = active_doc.do_export(src_path,select);
        // documents
        Com::CoUninitialize();
    }
}
// 创建文档
pub fn document_open(src:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app_res = Application::new(&app_ver);
        if app_res.is_none() {
            return;
        }
        let app = app_res.unwrap();
        let docs = app.get_documents();
        let res = docs.Open(src);
        // let sum  =selects.get_Count();
        println!("print task sum:{:?}",res);
        // let select =  selects.Item(0).unwrap();
        // let hr = active_doc.do_export(src_path,select);
        // documents
        Com::CoUninitialize();
    }
}


// 文件保存
pub fn document_save(src_path:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app_res = Application::new(&app_ver);
        if app_res.is_none() {
            return;
        }
        let app = app_res.unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let hr = active_doc.do_saveas(src_path);
        // documents
        Com::CoUninitialize();
    }
}

// 导出预览图
pub fn document_export_preview(){

}

// 导出预览图
pub fn document_export_cover(src_path:&str){
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_ver = "25".to_string();
        let app_res = Application::new(&app_ver);
        if app_res.is_none() {
            return;
        }
        let app = app_res.unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let name = active_doc.get_Name();
        println!("{:?}",name);
        let selects = active_doc.get_SelectionSets().unwrap();
        let sum  =selects.get_Count();
        println!("sum:{:?}",sum);
        // let act_select = active_doc.get_ActiveSelectionSet().unwrap();
        // let act_sum  =act_select.get_Count();
        // println!("sum:{:?}",act_sum);
        let select_item =  selects.Item(0).unwrap();
        let hr = active_doc.do_export(src_path,select_item);
        // documents
        Com::CoUninitialize();
    }
}

// 导出预览图
pub fn get_document_name(ver:&str)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app = Application::new(&ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let name = active_doc.get_Name();
        // documents
        return name;
        Com::CoUninitialize();
    }
}

// 导出预览图
pub fn get_document_path(ver:&str)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app = Application::new(&ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let path = active_doc.get_Path();
        // documents
        return path;
        Com::CoUninitialize();
    }
}

// 导出预览图
pub fn get_document_fullpath(ver:&str)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app = Application::new(&ver).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let path = active_doc.get_FullName();
        // documents
        return path;
        Com::CoUninitialize();
    }
}
