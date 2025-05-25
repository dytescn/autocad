use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadViewports{
    disp:ComObject
}

impl IAcadViewports{
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
    fn DeleteConfiguration(){

    }
}

pub struct IAcadViewport{
    disp:ComObject
}

impl IAcadViewport{
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
    fn get_GridOn(){
        
    }
    fn put_GridOn(){
        
    }
    fn get_OrthoOn(){
        
    }
    fn put_OrthoOn(){
        
    }
    fn get_SnapBasePoint(){
        
    }
    fn put_SnapBasePoint(){
        
    }
    fn get_SnapOn(){
        
    }
    fn put_SnapOn(){
        
    }
    fn get_SnapRotationAngle(){
        
    }
    fn put_SnapRotationAngle(){
        
    }
    fn get_UCSIconOn(){
        
    }
    fn put_UCSIconOn(){
        
    }
    fn get_UCSIconAtOrigin(){
        
    }
    fn put_UCSIconAtOrigin(){
        
    }
    fn get_LowerLeftCorner(){
        
    }
    fn get_UpperRightCorner(){
        
    }
    fn Split(){
        
    }
    fn GetGridSpacing(){
        
    }
    fn SetGridSpacing(){
        
    }
    fn GetSnapSpacing(){
        
    }
    fn SetSnapSpacing(){
        
    }
    fn SetView(){
        
    }
    fn get_ArcSmoothness(){
        
    }
    fn put_ArcSmoothness(){

    }
}