use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadGroups {
    disp:ComObject
}

impl IAcadGroups{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index: u8)->Option<IAcadGroup>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadGroup::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn get_Count(&self)->u8{
        let info = self.disp.get_property("count").log_error("got err").unwrap();
        info.to_u8().unwrap() 
    }
    fn get__NewEnum(){

    }
    fn Add(){

    }
}


pub struct IAcadGroup {
    disp:ComObject
}

impl IAcadGroup{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn Item(){

    }
    fn get__NewEnum(){

    }
    fn get_Count(){

    }
    fn put_TrueColor(){

    }
    fn put_Layer(){

    }
    fn put_Linetype(){

    }
    fn put_LinetypeScale(){

    }
    fn put_Visible(){

    }
    fn Highlight(){

    }
    fn put_PlotStyleName(){

    }
    fn put_Lineweight(){

    }
    pub fn get_Name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn put_Name(){

    }
    fn AppendItems(){

    }
    fn RemoveItems(){

    }
    fn Update(){

    }
    fn put_Material(){

    }
    fn put_color(){

    }
}