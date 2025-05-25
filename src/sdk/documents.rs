use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::document::IAcadDocument;

pub struct IAcadDocuments {
    disp:ComObject
}

impl IAcadDocuments{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index:u8)->Option<IAcadDocument>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadDocument::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    fn get__NewEnum(){

    }
    pub fn get_Count(&self)->u8{
        let info_res = self.disp.get_property("count");
        if info_res.is_err() {
            return 0;
        }
        let info = info_res.unwrap();
        info.to_u8().unwrap()
    }
    fn get_Application(){

    }
    pub fn Add(&self){
        let mut opts = Vec::new();
        let tpl: VARIANT = VARIANT::null();
        opts.push(tpl);
        let hres = self.disp.invoke_method("add", opts).unwrap();
    }
    // 打开文件
    pub fn Open(&self,src:&str)->Option<()>{
        let mut opts = Vec::new();
        let vart_src: VARIANT = VARIANT::from_str(src);
        opts.push(vart_src);
        let hres = self.disp.invoke_method("open", opts);
        if hres.is_err() {
            return None;
        }
            return  Some(());
    }
    fn Close(&self){

    }
}