
use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use crate::sdk::types::*;
use crate::sdk::document::IAcadDocument;

use super::documents::IAcadDocuments;
use super::preferences::IAcadPreferences;
use windows::core::GUID;

pub struct Application {
    disp:ComObject
}

impl Application{
    pub fn new(ver:&str)-> Option<Self> {
       let id_name = "autocad.application"; // Can't use + with two &str
       let res_data=  ComObject::new_from_app(&id_name);
       match res_data {
           Some(data)=>{
                return Some(Self{
                    disp:data
                });
                
           }
           None=>{
                return None;
           }
       }
    }
    pub fn new_from_clisd(ver:&str)-> Option<Self> {
        let id_name = "autocad.application"; // Can't use + with two &str
        let res_data=  ComObject::new_from_name(&id_name,GUID::zeroed());
        match res_data {
            Some(data)=>{
                 return Some(Self{
                     disp:data
                 });
                 
            }
            None=>{
                 return None;
            }
        }
     }

    pub fn get_visible(&self)->bool{
        let app_name = self.disp.get_property("visible").log_error("got err").unwrap();
        app_name.to_bool().log_error("got error").unwrap()        
    }
    fn put_visible(){
        
    }
    pub fn get_name(&self)->String{
        let app_name = self.disp.get_property("name").log_error("got err").unwrap();
        app_name.to_string().log_error("got error").unwrap()
    }
    fn get_caption(){
        
    }
    fn get_application(){
        
    }
    pub fn get_activedocument(&self)->Option<IAcadDocument>{
        let doc_res = self.disp.get_property("activedocument");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let acad_doc = IAcadDocument::new(disp.clone());
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
    fn put_activedocument(){
        
    }
    pub fn get_fullname(&self)->String{
        let info = self.disp.get_property("fullname").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn get_height(){
        
    }
    fn put_height(){
        
    }
    fn get_window_left(){
        
    }
    fn put_window_left(){
        
    }
    pub fn get_path(&self)->String{
        let info = self.disp.get_property("path").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn get_locale_id(){
        
    }
    fn get_window_top(){
        
    }
    fn put_windowtop(){
        
    }
    pub fn get_version(&self)->String{
        let app_vers = self.disp.get_property("version").log_error("got err").unwrap();
        app_vers.to_string().log_error("got error").unwrap()
    }
    pub fn get_width(&self)->u32{
        let app_width = self.disp.get_property("width").log_error("got err").unwrap();

        app_width.to_u32().log_error("got error").unwrap()
    }
    fn put_width(){
        
    }
    pub fn get_preferences(&self)->Option<IAcadPreferences>{
        let info_res = self.disp.get_property("preferences");
        match info_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let acad_doc = IAcadPreferences::new(disp.clone());
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
    fn get_statusid(){
        
    }
    fn list_arx(){
        
    }
    fn LoadArx(){
        
    }
    fn get_interface_object(){
        
    }
    fn unload_arx(){
        
    }
    fn update(){
        
    }
    pub fn do_quit(&self){
        let app_vers = self.disp.invoke_method("quit", Vec::new()).expect("got err");
    }
    fn zoom(){
        
    }
    fn get_vbe(){
        
    }
    fn get_menugroups(){
        
    }
    fn get_menubar(){
        
    }
    fn load_dvb(){
        
    }
    fn unload_dvb(){
        
    }
    pub fn get_documents(&self)->IAcadDocuments{
        let doc = self.disp.get_property("documents").log_error("got error").unwrap();
        let disp = doc.to_idispatch().expect("change error").clone();
        IAcadDocuments::new(disp)
    }
    fn eval(){
        
    }
    fn get_windowstate(){
        
    }
    fn put_windowstate(){
        
    }
    fn run_macro(){
        
    }
    fn zoom_extents(){
        
    }
    pub fn zoom_all(&self){
        
    }
    pub fn zoom_center(&self){
        
    }
    pub fn zoom_scaled(&self){
        
    }
    pub fn zoom_window(&self){
        
    }
    fn zoom_pick_window(&self){
        
    }
    fn get_acad_state(){
        
    }
    fn zoom_previous(){
        
    }
    fn get_hwnd(&self){

    }

}