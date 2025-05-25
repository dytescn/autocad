use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadPlotConfiguration {
    disp:ComObject
}

impl IAcadPlotConfiguration{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    // 打印配置名称
    pub fn get_Name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn put_Name(&self,name:&str)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_str(name);
        opts.push(vart_status);
        let info = self.disp.set_property("name",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 配置名
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
    // 打印媒介尺寸
    /*
    *  1. 16K_(15360.00_x_8640.00_Pixels)
    *  2. 4K_UHD_(3840.00_x_2160.00_Pixels)
    *  3. 5K_(5120.00_x_2880.00_Pixels)
    *  4. 8K_UHD_(7680.00_x_4320.00_Pixels)
    *  5. DCI_4K_(4096.00_x_2160.00_Pixels)
    *  6. FHD_(1920.00_x_1080.00_Pixels)
    *  7. HD_(1280.00_x_1080.00_Pixels)
    * 8. QHD_(2560.00_x_1440.00_Pixels)
    * 9. QHD+_(3200.00_x_1800.00_Pixels)
    */
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
    // 纸张单位
    // acInches = 0, //英寸
    // acMillimeters = 1, // 毫米
    // acPixels = 2,    // 像素
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
    // 打印旋转
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
    // 打印是否居中
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
    // 打印隐藏
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
    // 打印方式
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
    // 打印预览
    pub fn get_ViewToPlot(&self)->String{
        let info = self.disp.get_property("viewtoplot").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_ViewToPlot(&self,typing:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_i32(typing);
        opts.push(vart_status);
        let info = self.disp.set_property("viewtoplot",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 自动缩放
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
    // 设置缩放
    pub fn get_StandardScale(&self)->i32{
        let info = self.disp.get_property("standardscale").log_error("got err").unwrap();
        info.to_i32().log_error("got error").unwrap()
    }
    pub fn put_StandardScale(&self,scale:i32)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_i32(scale);
        opts.push(vart_status);
        let info = self.disp.set_property("standardscale",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 用户缩放
    pub fn GetCustomScale(&self){

    }
    pub fn SetCustomScale(&self){

    }
    // 缩放线宽
    pub fn get_ScaleLineweights(&self)->bool{
        let info = self.disp.get_property("scalelineweights").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_ScaleLineweights(&self,status:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(status);
        opts.push(vart_status);
        let info = self.disp.set_property("scalelineweights",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 打印线宽
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
    // 获取第一个打印视窗口
    pub fn get_PlotViewportsFirst(&self)->bool{
        let info = self.disp.get_property("plotviewportsfirst").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    pub fn put_PlotViewportsFirst(&self,status:bool)->std::fmt::Result{
        let mut opts: Vec<VARIANT> = Vec::new();
        let vart_status: VARIANT = <VARIANT as VariantExt>::from_bool(status);
        opts.push(vart_status);
        let info = self.disp.set_property("plotviewportsfirst",opts);
        if info.is_err() {
            let err = std::fmt::Error::default();
            return Err(err);
        }
        return Ok(());
    }
    // 获取样式列表
    pub fn get_StyleSheet(&self)->String{
        let info = self.disp.get_property("stylesheet").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    pub fn put_StyleSheet(&self,style:String)->std::fmt::Result{
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
    // 纸边距
    pub fn get_PaperMargins(&self){
        let info = self.disp.get_property("papermargins").log_error("got err").unwrap();
        // let arr = info.to
        // dbg!(info)
        // print!("{:?}",dbg!(info));
        // info.to_string().log_error("got error").unwrap()
    }
    pub fn getPaperSize(&self){
        let info = self.disp.get_property("papersize").log_error("got err").unwrap();
        // info.to_array();
    }

    pub fn get_PlotOrigin(&self)->i32{
        let infos = self.disp.get_property("plotorigin").log_error("got err").unwrap();
        return 0;
    }
    pub fn put_PlotOrigin(&self){

    }
    // 获取打印窗口
    pub fn GetWindowToPlot(&self){

    }
    pub fn SetWindowToPlot(&self){

    }
    // 获取打印样式
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
    // 获取模型类型
    pub fn get_ModelType(&self)->bool{
        let info = self.disp.get_property("modeltype").log_error("got err").unwrap();
        info.to_bool().log_error("got error").unwrap()
    }
    // 复制打印配置
    pub fn CopyFrom(&self){

    }
    // 获取媒体类型列表
    pub fn GetCanonicalMediaNames(&self){

    }
    pub fn GetPlotDeviceNames(&self){
        // let info = self.disp.get_property("configname").log_error("got err").unwrap();
        // info.to_string().log_error("got error").unwrap()
    }
    
    // 打印样式表
    pub fn GetPlotStyleTableNames(&self)->String{
        let info = self.disp.get_property("PlotStyleTableNames").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    
    // 刷新设备
    pub fn RefreshPlotDeviceInfo(&self){

    }

    // 获取本地媒体名
    pub fn GetLocaleMediaName(&self){
        let info = self.disp.get_property("configname").log_error("got err").unwrap();
        // info.to_string().log_error("got error").unwrap()
    }
}