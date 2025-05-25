use std::ptr::null;

use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::entity::IAcadEntity;


pub struct IAcadSelection{
    disp:ComObject
}

impl IAcadSelection{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){

    }
    fn put_PickFirst(){

    }
    fn get_PickFirst(){

    }
    fn put_PickAdd(){

    }
    fn get_PickAdd(){

    }
    fn put_PickDrag(){

    }
    fn get_PickDrag(){

    }
    fn put_PickAuto(){

    }
    fn get_PickAuto(){

    }
    fn put_PickBoxSize(){

    }
    fn get_PickBoxSize(){

    }
    fn put_DisplayGrips(){

    }
    fn get_DisplayGrips(){

    }
    fn put_DisplayGripsWithinBlocks(){

    }
    fn get_DisplayGripsWithinBlocks(){

    }
    fn put_GripColorSelected(){

    }
    fn get_GripColorSelected(){

    }
    fn put_GripColorUnselected(){

    }
    fn get_GripColorUnselected(){

    }
    fn put_GripSize(){

    }
    fn get_GripSize(){

    }
    fn put_PickGroup(){

    }
    fn get_PickGroup(){

    }
    pub fn get_opt(&self)->VARIANT{
        self.disp.get_variant().expect("got variant err:")
    }
}
