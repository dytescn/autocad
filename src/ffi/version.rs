use std::num::ParseIntError;
use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::core;

// 获取coreldraw
pub fn get_version_info(ver:&str)->Option<u32>{
    unsafe{
        // 初始化系统
        let com_res = Com::CoInitialize(None);
        if com_res.is_err() {
            println!("com init error");
            return None;
        }
        let id_name = "autocad.application.".to_string() + ver; // Can't use + with two &str
        let lpsz = core::HSTRING::from(id_name);
        let rclsid_res = Com::CLSIDFromProgID(&lpsz);
        if rclsid_res.is_err(){
            return None;
        }
        let res_ver_n = ver.parse();
        if res_ver_n.is_err() {
            return None;
        }
        let ver_n = res_ver_n.unwrap();
        Com::CoUninitialize();
        return Some(ver_n);
    }
}

// 判断当前已安装的coreldraw版本
pub fn get_version_list()->Vec<u32>{
    let mut ver_data:Vec<u32> = Vec::new();
    for i in 15..28  {
        let ver = i.to_string();
        let res = get_version_info(&ver);
        match res {
            Some(ver_name)=>{
                ver_data.push(ver_name);
            }
            None=>{
                println!("...{}",i);
            }
        }
    }
    return ver_data;
}
