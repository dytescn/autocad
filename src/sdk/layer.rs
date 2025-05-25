use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadLayers{
    disp:ComObject
}

impl IAcadLayers{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index:u8)->Option<IAcadLayer>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadLayer::new(disp.clone());
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
    fn GenerateUsageData(){

    }
}


pub struct IAcadLayer {
    disp:ComObject
}

impl IAcadLayer{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_color(){

    }
    fn put_color(){

    }
    fn get_TrueColor(){

    }
    fn put_TrueColor(){

    }
    fn get_Freeze(){

    }
    fn put_Freeze(){

    }
    fn get_LayerOn(){

    }
    fn put_LayerOn(){

    }
    fn get_Linetype(){

    }
    fn put_Linetype(){

    }
    fn get_Lock(){

    }
    fn put_Lock(){

    }
    fn get_Name(){

    }
    fn put_Name(){

    }
    fn get_Plottable(){

    }
    fn put_Plottable(){

    }
    fn get_ViewportDefault(){

    }
    fn put_ViewportDefault(){

    }

    pub fn get_PlotStyleName(&self)->String{
        let info = self.disp.get_property("plotstylename").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn put_PlotStyleName(){

    }
    fn get_Lineweight(){

    }
    fn put_Lineweight(){

    }
    fn get_Description(){

    }
    fn put_Description(){

    }
    fn get_Used(){

    }
    fn get_Material(){

    }
    fn put_Material(){

    }
    pub fn get_vart(&self)->VARIANT{
        self.disp.get_variant().log_error("change error").unwrap()
    }
}