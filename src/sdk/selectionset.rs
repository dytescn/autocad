use std::ptr::null;

use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::entity::IAcadEntity;


pub struct IAcadSelectionSet {
    disp:ComObject
}

impl IAcadSelectionSet{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index: u8)->Option<IAcadEntity>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc: IAcadEntity = IAcadEntity::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn get_Count(&self)->i32{
        let info = self.disp.get_property("count").log_error("got err").unwrap();
        info.to_i32().unwrap() 
    }
    fn get__NewEnum(){

    }
    pub fn get_Name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();
        info.to_string().unwrap() 
    }
    pub fn Highlight(&self){

    }
    fn Erase(){

    }
    fn Update(){

    }
    fn get_Application(){

    }
    fn AddItems(){

    }
    fn RemoveItems(){

    }
    fn Clear(){

    }
    fn Select(){

    }
    fn SelectAtPoint(){

    }
    fn SelectByPolygon(){

    }
    fn SelectOnScreen(){

    }
    fn Delete(){

    }
    pub fn get_opt(&self)->VARIANT{
        self.disp.get_variant().expect("got variant err:")
    }
}