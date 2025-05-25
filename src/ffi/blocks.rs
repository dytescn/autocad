use dyteslogs::logs::LogError;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use crate::sdk::application::Application;
// 获取coreldraw信息
pub fn blocks_info(cur_version:String){
    unsafe{
        // 初始化系统
        let res  = Com::CoInitialize(None);
        if res.is_err() {
            return;
        }
        let mut sum = 0;
        let app = Application::new(&cur_version).unwrap();
        let active_doc = app.get_activedocument().unwrap();
        let actLayout = active_doc.get_ActiveLayout().unwrap();
        let allblock = actLayout.get_Block().unwrap();
        print!("................");
        let count = allblock.get_Count();
        for i in 0..count  {
            // println!("----{:?}----",i);
            let enty = allblock.Item(i).unwrap();
            let enty_type = enty.get_EntityType();
            let enty_name = enty.get_EntityName();
            if enty_type==7 {
                sum = sum +1;
                println!("type::{:?},name::{:?}",enty_type,enty_name);
            }
        }
        println!("sum:::{:?}",sum);
        // documents
        Com::CoUninitialize();
    }
}