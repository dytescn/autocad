use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use windows::Win32::System::Variant;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::block::IAcadBlock;

pub struct IAcadLayouts{
    disp:ComObject
}

impl IAcadLayouts{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index: u8)->Option<IAcadLayout>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadLayout::new(disp.clone());
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
    fn Add(){

    }
}


pub struct IAcadLayout {
    disp:ComObject
}

impl IAcadLayout{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn get_Block(&self)->Option<IAcadBlock>{
        let hres = self.disp.get_property("block").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadBlock::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    pub fn get_TabOrder(&self)->u8{
        let info = self.disp.get_property("taborder").log_error("got err").unwrap();
        info.to_u8().log_error("got error").unwrap()
    }
    pub fn put_TabOrder(&self,index:i32){
        let mut params = Vec::new();
        let taborder = <VARIANT as VariantExt>::from_i32(index);
        params.push(taborder);
        self.disp.set_property("taborder",params).expect("got error");   
    }
    pub fn get_ConfigName(&self)->String{
        let info = self.disp.get_property("configname").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_ConfigName(&self,name:&str)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_str(name);
        opts.push(vart_status);
        let info = self.disp.set_property("configname",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_CenterPlot(&self)->bool{
        let info = self.disp.get_property("centerplot").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_CenterPlot(&self,center:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(center);
        opts.push(vart_status);
        let info = self.disp.set_property("centerplot",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_StyleSheet(&self)->String{
        let info = self.disp.get_property("stylesheet").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_StyleSheet(&self,style:&str)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_str(style);
        opts.push(vart_status);
        let info = self.disp.set_property("stylesheet",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_PlotWithLineweights(&self)->bool{
        let info = self.disp.get_property("plotwithlineweights").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_PlotWithLineweights(&self,status:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(status);
        opts.push(vart_status);
        let info = self.disp.set_property("plotwithlineweights",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_CanonicalMediaName(&self)->String{
        let info = self.disp.get_property("canonicalmedianame").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_CanonicalMediaName(&self,name:&str)->std::fmt::Result{        
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_str(name);
        opts.push(vart_status);
        let info = self.disp.set_property("canonicalmedianame",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_PlotRotation(&self)->i32{
        let info = self.disp.get_property("plotrotation").log_error("got err").unwrap();
        info.to_i32().log_error("got error").unwrap()
    }
    pub fn put_PlotRotation(&self,rotat:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_i32(rotat);
        opts.push(vart_status);
        let info = self.disp.set_property("plotrotation",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_PlotWithPlotStyles(&self)->bool{
        let info = self.disp.get_property("plotwithplotstyles").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_PlotWithPlotStyles(&self,status:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(status);
        opts.push(vart_status);
        let info = self.disp.set_property("plotwithplotstyles",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn get_PlotHidden(&self)->bool{
        let info = self.disp.get_property("plothidden").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_PlotHidden(&self,hidden:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(hidden);
        opts.push(vart_status);
        let info = self.disp.set_property("plothidden",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // PlotType (enum类型):
    // 0  acDisplay: 按显示的内容打印.   显示
    // 1 acExtents: 按当前选定空间范围内的所有内容打印.   范围
    // 2 acLimits: 打印当前空间范围内的所有内容.  范围  ？？
    // 3 acView: 打印由 ViewToPlot 属性命名的视图.
    // 4 acWindow: 打印由 SetWindowToPlot 方法指定的窗口中的所有内容.  ******
    // 5  acLayout: 打印位于指定纸张尺寸边缘的所有内容，原点从 0,0 坐标计算。 
    pub fn get_PlotType(&self)->i32{
        let info = self.disp.get_property("plottype").log_error("got err").unwrap();
        info.to_i32().log_error("got error").unwrap()
    }
    pub fn put_PlotType(&self,typing:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_i32(typing);
        opts.push(vart_status);
        let info = self.disp.set_property("plottype",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    pub fn GetWindowToPlot(&self){
            let mut opts: Vec<VARIANT> = Vec::new();
            let mut vart_left: VARIANT = <VARIANT as VariantExt>::null();
            opts.push(vart_left);
            let mut vart_right: VARIANT = <VARIANT as VariantExt>::null();
            opts.push(vart_right);
            let info = self.disp.invoke_method("GetWindowToPlot",opts).log_error("got err").unwrap();
            let data_var = info.to_vart_array().unwrap();
            let sum = data_var.len();
            println!("sum:{:?}",sum);
            // let data1 =  info.to_vart_array().unwrap();
            // println!("{:?}",data1);
            // let data2 =  info[1].to_f64_array().unwrap();
            // println!("{:?}",data2);
        
            // let vart_left_arr = info.to_vart_array().unwrap();
            // let left = vart_left_arr[0].to_f64_array().unwrap();
            // println!("{:?}",left);
            // println!("{:?}",vart_left_arr.len());
        // let arr = info.to
        // dbg!(info)
        // print!("{:?}",dbg!(info));
        // info.to_string().log_error("got error").unwrap()
        // info.to_array();
    }
    
    pub fn SetWindowToPlot(&self,LowerLeft: Vec<f64>, UpperRight:Vec<f64>)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_left: VARIANT = <VARIANT as VariantExt>::from_vec_f64(LowerLeft);
        opts.push(vart_left);
        let vart_right: VARIANT = <VARIANT as VariantExt>::from_vec_f64(UpperRight);
        opts.push(vart_right);
        let info = self.disp.invoke_method("SetWindowToPlot",opts);
        if info.is_err() {
            println!("is err:---{:?}",info.err());
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
        
    pub fn SetWindowToPlot_vec_vart(&self,opts: Vec<VARIANT>)->std::fmt::Result{
        let info = self.disp.invoke_method("SetWindowToPlot",opts);
        if info.is_err() {
            println!("is err:---{:?}",info.err());
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 获取打印尺寸
    pub fn getPaperSize(&self){
        let info = self.disp.get_property("getpapersize").log_error("got err").unwrap();
        // info.to_array();
    }
    pub fn get_PlotOrigin(&self)->Vec<f64>{
        let infos = self.disp.get_property("plotorigin").log_error("got err").unwrap();
        infos.to_f64_array().unwrap()
    }
    pub fn put_PlotOrigin(&self,data:Vec<f64>){
        let mut opts: Vec<VARIANT> = Vec::new();
        let origin: VARIANT = <VARIANT as VariantExt>::from_vec_f64(data);
        opts.push(origin);        
        let info = self.disp.set_property("plotorigin",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            // return Err(err);
            println!("---{:?}---",info.err());
        }
    }
    pub fn get_PaperMargins(&self){
        let info = self.disp.get_property("papermargins").log_error("got err").unwrap();
        // let arr = info.to
        // dbg!(info)
        // print!("{:?}",dbg!(info));
        // info.to_string().log_error("got error").unwrap()
        // info.to_array();
    }
    pub fn get_UseStandardScale(&self)->bool{
        let info = self.disp.get_property("usestandardscale").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_UseStandardScale(&self,scale:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(scale);
        opts.push(vart_status);
        let info = self.disp.set_property("usestandardscale",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }

    pub fn get_PaperUnits(&self)->i32{
        let info = self.disp.get_property("paperunits").log_error("got err").unwrap();
        info.to_i32().log_error("got error").unwrap()
    }
    pub fn put_PaperUnits(&self,unit:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_i32(unit);
        opts.push(vart_status);
        let info = self.disp.set_property("paperunits",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }

    // 打印预览框
    pub fn get_PlotViewportBorders(&self)->bool{
        let info = self.disp.get_property("plotviewportborders").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_PlotViewportBorders(&self,border:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(border);
        opts.push(vart_status);
        let info = self.disp.set_property("plotviewportborders",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
        // 打印样式
        pub fn get_ShowPlotStyles(&self)->bool{
            let info = self.disp.get_property("showplotstyles").log_error("got err").unwrap();
            info.to_bool().log_error("got error").unwrap()
        }
        pub fn put_ShowPlotStyles(&self,style:bool)->std::fmt::Result{
            let mut opts: Vec<VARIANT> = Vec::new();
            let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(style);
            opts.push(vart_status);
            let info = self.disp.set_property("showplotstyles",opts);
            if info.is_err() {
                let err = std::fmt::Error::default();
                return Err(err);
            }
            return Ok(()); 
        }

}