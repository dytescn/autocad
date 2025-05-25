use std::alloc::Layout;
use std::ptr::null;

use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::database::IAcadDatabase;
use super::layouts::IAcadLayout;
use super::plot::IAcadPlot;
use super::selectionset::IAcadSelectionSet;
use super::selectionsets::IAcadSelectionSets;


pub struct IAcadDocument {
    disp:ComObject
}

impl IAcadDocument{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn put_name(&self,name:&str){
        let mut vart_vec = Vec::new();
        let vart_name = <VARIANT as VariantExt>::from_str(name);
        vart_vec.push(vart_name);
        self.disp.set_property("name", vart_vec);
    }

    pub fn get_Plot(&self)->Option<IAcadPlot>{
        let hres = self.disp.get_property("plot").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadPlot::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    // 获取修改的层
    fn get_ActiveLayer(){
        
    }
    // 修改当前选中层
    fn put_ActiveLayer(){
        
    }
    // 获取选中线的类型
    fn get_ActiveLinetype(){
        
    }
    // 修改选中线的类型 
    fn put_ActiveLinetype(){
        
    }
    fn get_ActiveDimStyle(){
        
    }
    fn put_ActiveDimStyle(){
        
    }
    fn get_ActiveTextStyle(){
        
    }
    fn put_ActiveTextStyle(){
        
    }
    fn get_ActiveUCS(){
        
    }
    fn put_ActiveUCS(){
        
    }
    fn get_ActiveViewport(){
        
    }
    fn put_ActiveViewport(){
        
    }
    fn get_ActivePViewport(){
        
    }
    fn put_ActivePViewport(){
        
    }
    fn get_ActiveSpace(){
        
    }
    fn put_ActiveSpace(){
        
    }
    pub fn get_SelectionSets(&self)->Option<IAcadSelectionSets>{
        let hres = self.disp.get_property("selectionsets").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadSelectionSets::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn get_ActiveSelectionSet(&self)->Option<IAcadSelectionSet>{
        let hres = self.disp.get_property("activeselectionset").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadSelectionSet::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn get_FullName(&self)->String{
        let info = self.disp.get_property("fullname").log_error("got err").unwrap();

        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_fullfilename(&self,src_path:&str)->Option<()>{
        let mut vart_vec = Vec::new();
        let vart_path = <VARIANT as VariantExt>::from_str(src_path);
        vart_vec.push(vart_path);
        let info = self.disp.set_property("fullname",vart_vec);
        if info.is_err() {
            return None;
        }
        return Some(());
    }
    pub fn get_Name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();

        info.to_string().log_error("got error").unwrap()
    }
    pub fn get_Path(&self)->String{
        let info = self.disp.get_property("path").log_error("got err").unwrap();

        info.to_string().log_error("got error").unwrap()
    }
    fn get_ObjectSnapMode(){
        
    }
    fn put_ObjectSnapMode(){
        
    }
    fn get_ReadOnly(){
        
    }
    fn get_Saved(){
        
    }
    fn get_MSpace(){
        
    }
    fn put_MSpace(){
        
    }
    fn get_Utility(){
        
    }
    fn Open(){
        
    }
    fn AuditInfo(){
        
    }
    fn Import(){
        
    }
    pub fn do_export(&self,src_path:&str,select:IAcadSelectionSet){
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(src_path);
        opts.push(vart_file);
        // 类型
        let vart_type = <VARIANT as VariantExt>::from_str("PDF");
        opts.push(vart_type);
        // 选择区域
        let select_vart =  select.get_opt();
        opts.push(select_vart);
        let hres = self.disp.invoke_method("export", opts);
        if hres.is_err() {
            println!("{:?}",hres.err());
        }       
    }
    fn New(){
        
    }
    pub fn do_save(&self)->Option<()>{
        let mut opts = Vec::new();
        let hres = self.disp.invoke_method("save", opts);
        if hres.is_err() {
            return None;
        }
        return Some(());
    }
    pub fn do_saveas(&self,path_str:&str)->Option<()>{
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(path_str);
        opts.push(vart_file);
        let hres = self.disp.invoke_method("saveas", opts);
        if hres.is_err() {
            return None;
        }
        return Some(());
    }
    pub fn do_saveas_dxf(&self,path_str:&str)->Option<()>{
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(path_str);
        opts.push(vart_file);
        
        let vart_file_type = <VARIANT as VariantExt>::from_i32(49);
        opts.push(vart_file_type);
        let hres = self.disp.invoke_method("saveas", opts);
        if hres.is_err() {
            println!("{:?}",hres.err());
            return None;
        }
        return Some(());
    }
    fn Wblock(){
        
    }
    fn PurgeAll(){
        
    }
    fn GetVariable(){
        
    }
    pub fn SetVariable(&self,able:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_able: VARIANT = <VARIANT as VariantExt>::from_i32(able);
        opts.push(vart_able);
        let info = self.disp.invoke_method("SetWindowToPlot",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    fn LoadShapeFile(){
        
    }
    fn Regen(){
        
    }
    fn get_PickfirstSelectionSet(){
        
    }
    fn get_Active(){
        
    }
    fn Activate(){
        
    }
    fn Close(){
        
    }
    fn put_WindowState(){
        
    }
    fn get_WindowState(){
        
    }
    fn put_Width(){
        
    }
    fn get_Width(){
        
    }
    fn put_Height(){
        
    }
    fn get_Height(){
        
    }
    fn put_ActiveLayout(){
        
    }
    pub fn get_ActiveLayout(&self)->Option<IAcadLayout>{
        let doc_res = self.disp.get_property("activelayout");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let acad_doc = IAcadLayout::new(disp.clone());
                        return Some(acad_doc);
                    }
                    Err(e)=>{
                        return None
                    }
                }
            }
            Err(err)=>{
                return None
            }
        }
    }
    fn SendCommand(){
        
    }
    fn get_HWND(){
        
    }
    fn get_WindowTitle(){
        
    }
    fn get_Application(){
        
    }
    pub fn get_Database(&self)->Option<IAcadDatabase>{
        let hres = self.disp.get_property("database").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadDatabase::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    fn StartUndoMark(){
        
    }
    fn EndUndoMark(){
        
    }
    fn get_ActiveMaterial(){
        
    }
    fn put_ActiveMaterial(){
        
    }
    fn PostCommand(){

    }
    pub fn do_close(&self,src_path:&str)->Option<()>{
        let mut opts = Vec::new();
        // 是否保存
        let vart_file = <VARIANT as VariantExt>::from_bool(true);
        opts.push(vart_file);
        // 文件地址
        let vart_file = <VARIANT as VariantExt>::from_str("src_path");
        opts.push(vart_file);
        let hres = self.disp.invoke_method("close", opts);
        if hres.is_err() {
            println!("{:?}",hres.err());
            return None;
        }
            return Some(());
    }

}