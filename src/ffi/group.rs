use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
// 获取coreldraw信息
pub fn do_group(cur_version:String)->bool{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = Application::new(&cur_version).unwrap();
        
        return true;
        Com::CoUninitialize();
    }
}