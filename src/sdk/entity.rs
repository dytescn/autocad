use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::layer::IAcadLayer;

pub struct IAcadEntity {
    disp:ComObject
}

impl IAcadEntity{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_TrueColor(){

    }
    fn put_TrueColor(){

    }
    // 
    pub fn get_Layer(&self)->Option<IAcadLayer>{
        let info = self.disp.get_property("layer").log_error("got err").unwrap();
        match info.to_idispatch(){
            Ok(disp)=>{
                let layer: IAcadLayer = IAcadLayer::new(disp.clone());
                return Some(layer);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn put_Layer(&self,layer:IAcadLayer)->bool{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_visiable: VARIANT = layer.get_vart();
        opts.push(vart_visiable);
        let info = self.disp.set_property("visible",opts);
        if info.is_err(){
            return false;
        }
        return  true;
    }
    fn get_Linetype(){

    }
    fn put_Linetype(){

    }
    fn get_LinetypeScale(){

    }
    fn put_LinetypeScale(){

    }
    pub fn get_Visible(&self)->bool{
        let info = self.disp.get_property("visible").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_Visible(&self,visiable:bool){
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_visiable: VARIANT = <VARIANT as VariantExt>::from_bool(visiable);
        opts.push(vart_visiable);
        let info = self.disp.set_property("visible",opts);
    }
    fn ArrayPolar(){

    }
    fn ArrayRectangular(){

    }
    fn Highlight(){

    }
    fn Copy(){

    }
    fn Move(){

    }
    fn Rotate(){

    }
    fn Rotate3D(){

    }
    fn Mirror(){

    }
    fn Mirror3D(){

    }
    fn ScaleEntity(){

    }
    fn TransformBy(){

    }
    fn Update(){

    }
    fn GetBoundingBox(){

    }
    fn IntersectWith(){

    }
    pub fn get_PlotStyleName(&self)->String{
        let info = self.disp.get_property("plotstylename").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    
    pub fn put_PlotStyleName(&self,name:&str)->bool{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_name: VARIANT = <VARIANT as VariantExt>::from_str(name);
        opts.push(vart_name);
        let info = self.disp.set_property("plotstylename",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return false;
        }
        return true;
    }
    fn get_Lineweight(){

    }
    fn put_Lineweight(){

    }
    fn get_EntityTransparency(){

    }
    fn put_EntityTransparency(){

    }
    fn get_Hyperlinks(){

    }
    fn get_Material(){

    }
    fn put_Material(){

    }
    pub fn get_EntityName(&self)->String{
        let info = self.disp.get_property("entityname").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn get_name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn get_EntityType(&self)->i32{
        let info = self.disp.get_property("entitytype");
        if info.is_err() {
            return 0;
        }
        let info_data = info.unwrap();
        info_data.to_i32().log_error("got error").unwrap()
    }
    fn get_color(){

    }
    fn put_color(){

    }
    pub fn get_Coordinates(&self)->VARIANT{
        let info = self.disp.get_property("Coordinates").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }
    fn put_Coordinates(&self){
        
    }
    pub fn get_Coordinate(&self)->VARIANT{
        let id = <VARIANT as VariantExt>::from_i32(1);
        let info = self.disp.get_property_by_vart("Coordinate",id).log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }
    fn put_Coordinate(&self){

    }
    pub fn get_StartPoint(&self)->VARIANT{
        let info = self.disp.get_property("startpoint").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }
    pub fn get_EndPoint(&self)->VARIANT{
        let info = self.disp.get_property("endpoint").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }
    pub fn get_InsertionPoint(&self)->VARIANT{
        let info = self.disp.get_property("InsertionPoint").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }

    pub fn get_Center(&self)->VARIANT{
        let info = self.disp.get_property("center").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        return info;
    }

    pub fn get_center_tt(&self)->Option<VARIANT>{
        let info = self.disp.get_property("center");
        if info.is_err(){
            return None;
        }
        return Some(info.unwrap());
    }
    pub fn get_insertionpoint(&self)->Option<VARIANT>{
        let info = self.disp.get_property("insertionpoint");
        if info.is_err(){
            return None;
        }
        return Some(info.unwrap());
    }
    pub fn get_HasAttributes(&self)->bool{
        let info = self.disp.get_property("HasAttributes");
        if info.is_err() {
            return false;
        }
        info.unwrap().to_bool().unwrap()
    }
    pub fn get_GetBoundingBox(&self)->Vec<VARIANT>{
        let mut params:[VARIANT;2] = [VARIANT::default(),VARIANT::default()];
        
        let info = self.disp.invoke_callback("GetBoundingBox", params.to_vec());
        if info.is_err() {
            println!("err:{:?}",info.err());
            return Vec::new();
        }
        let info_vart = info.unwrap();
        // let vart = info_vart.to_vart_array().unwrap();
        return info_vart;
    }
    
    
    pub fn get_Width(&self)->i64{
        let info = self.disp.get_property("width").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        // return info;
        info.to_i64().unwrap()
    }
    pub fn get_Height(&self)->i64{
        let info = self.disp.get_property("height").log_error("got err").unwrap();
        // let count =  Variant::VariantGetElementCount(self) as i32;
        // println!("-----{:?}-----",count);
        // return info;
        info.to_i64().unwrap()
    }
}