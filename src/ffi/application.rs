use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
use crate::sdk::document::IAcadDocument;

// 获取coreldraw信息
pub fn do_open(cdr_src:String,cur_version:String)->Option<()>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        if app_res.is_none() {
            return None;
        }
        let app = app_res.unwrap();
        let docs = app.get_documents();
        let res = docs.Open(&cdr_src);
        if res.is_none() {
            return None;
        }
        return Some(());
        Com::CoUninitialize();
    }
}

// 获取coreldraw信息
pub fn do_quit(cur_version:String)->bool{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        match app_res {
            Some(app)=>{
                app.do_quit();
                return true;
            }
            None=>{
                // return "aa".to_string();
                return false;
            }
        }
        Com::CoUninitialize();
    }
}


// 获取coreldraw信息
pub fn get_version(cur_version:String)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        match app_res {
            Some(app)=>{
                let ver = app.get_visible();
                // return  "aaa";
                println!("---{:?}---",ver);
                return "aaaaaa".to_string();
            }
            None=>{
                return "aa".to_string();
            }
        }
        Com::CoUninitialize();
    }
}

// 获取coreldraw信息
pub fn get_app_info(cur_version:String)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        if app_res.is_none() {
            return "".to_string();
        }
        let app = app_res.unwrap();
        let fullname = app.get_fullname();
        println!("{:?}",fullname);
        return  "".to_string();
        Com::CoUninitialize();
    }
}

// 获取coreldraw信息
pub fn get_cad_execute_path(cur_version:String)->String{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        if app_res.is_none() {
            return "".to_string();
        }
        let app = app_res.unwrap();
        let fullname = app.get_fullname();
        return fullname;
        Com::CoUninitialize();
    }
}

// 获取coreldraw信息
pub fn get_documents_sum(cur_version:String)->Option<usize>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version);
        if app_res.is_none() {
            return None;
        }
        let app = app_res.unwrap();
        let docs = app.get_documents();
        let sum =  docs.get_Count();
        Com::CoUninitialize();
        return Some(sum as usize);
    }
}
