use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadViews{
    disp:ComObject
}

impl IAcadViews{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn Item(){
        
    }
    fn get_Count(){
        
    }
    fn get__NewEnum(){
        
    }
    fn Add(){

    }
}


pub struct IAcadView {
    disp:ComObject
}

impl IAcadView{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Center(){

    }
    fn put_Center(){

    }
    fn get_Height(){

    }
    fn put_Height(){

    }
    fn get_Width(){

    }
    fn put_Width(){

    }
    fn get_Target(){

    }
    fn put_Target(){

    }
    fn get_Direction(){

    }
    fn put_Direction(){

    }
    fn get_Name(){

    }
    fn put_Name(){

    }
    fn get_CategoryName(){

    }
    fn put_CategoryName(){

    }
    fn get_LayoutId(){

    }
    fn put_LayoutId(){

    }
    fn get_LayerState(){

    }
    fn put_LayerState(){

    }
    fn get_HasVpAssociation(){

    }
    fn put_HasVpAssociation(){

    }
}