use std::fmt::Result;

use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::plotconfiguration::IAcadPlotConfiguration;

pub struct IAcadPlot {
    disp:ComObject
}

impl IAcadPlot{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn get_Application(){

    }
    fn get_QuietErrorMode(){

    }
    fn put_QuietErrorMode(){

    }
    pub fn get_NumberOfCopies(&self)->i32{
        let info = self.disp.get_property("numberofcopies").log_error("got err").unwrap();
        info.to_i32().log_error("got error").unwrap()
    }
    pub fn put_NumberOfCopies(&self,num:i32)->Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_visiable: VARIANT = <VARIANT as VariantExt>::from_i32(num);
        opts.push(vart_visiable);
        let info = self.disp.set_property("numberofcopies",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_BatchPlotProgress(&self)->bool{
        let info = self.disp.get_property("batchplotprogress").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_BatchPlotProgress(&self,status:bool)->Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(status);
        opts.push(vart_status);
        let info = self.disp.set_property("batchplotprogress",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn DisplayPlotPreview(&self)->bool{
        let mut opts = Vec::new();
        let cfg_vart =  <VARIANT as VariantExt>::from_i32(1);
        opts.push(cfg_vart);
        let hres = self.disp.invoke_method("displayplotpreview", opts);
        if hres.is_err() {
            println!("{:?}",hres.err());
            return false;
        } 
        return true;
    }
    pub fn PlotToFile(&self,path:&str,name:&str)->Option<()>{
        let mut opts = Vec::new();
        // 目录设置
        let vart_path = <VARIANT as VariantExt>::from_str(path);
        opts.push(vart_path);
        // let hres_plot = VARIANT::from_bool(true);
        // opts.push(hres_plot);
        let hres = self.disp.invoke_method("plottofile", opts);
        if hres.is_err() {
            // println!("plot file errr::::{:?}",hres.err());
            return None;
        } 
        return Some(());
    }
    fn PlotToDevice(){

    }
    fn SetLayoutsToPlot(){

    }
    // 开始批处理模式
    pub fn StartBatchMode(&self,count:usize)->Option<()>{
        let mut opts = Vec::new();
        // 目录设置
        let vart_count = <VARIANT as VariantExt>::from_i32(count as i32);
        opts.push(vart_count);
        // let hres_plot = VARIANT::from_bool(true);
        // opts.push(hres_plot);
        let hres = self.disp.invoke_method("StartBatchMode", opts);
        if hres.is_err() {
            println!("{:?}",hres.err());
            return None;
        } 
        return Some(());
    }
}