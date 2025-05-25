use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;

use super::entity::IAcadEntity;

pub struct IAcadBlocks {
    disp:ComObject
}

impl IAcadBlocks{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index: u8)->Option<IAcadBlock>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u8(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc: IAcadBlock = IAcadBlock::new(disp.clone());
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
    fn get__NewEnum(){}
    fn Add(){}
}

pub struct IAcadBlock {
    disp:ComObject
}

impl IAcadBlock{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn Item(&self,index: u32)->Option<IAcadEntity>{
        let mut opts = Vec::new();
        let item_index: VARIANT = <VARIANT as VariantExt>::from_u32(index);
        opts.push(item_index);
        let hres = self.disp.invoke_method("item", opts).unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc: IAcadEntity = IAcadEntity::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    fn get__NewEnum(){}
    pub fn get_Count(&self)->u32{
        let info = self.disp.get_property("count").log_error("got err").unwrap();
        info.to_u32().log_error("got error").unwrap()  
    }
    pub fn get_Name(&self)->String{
        let info = self.disp.get_property("name").log_error("got err").unwrap();
        info.to_string().log_error("got error").unwrap()
    }
    fn put_Name(){}
    fn get_Origin(){}
    fn put_Origin(){}
    fn AddCustomObject(){}
    fn Add3DFace(){}
    fn Add3DMesh(){}
    fn Add3DPoly(){}
    fn AddArc(){}
    fn AddAttribute(){}
    fn AddBox(){}
    fn AddCircle(){}
    fn AddCone(){}
    fn AddCylinder(){}
    fn AddDimAligned(){}
    fn AddDimAngular(){}
    fn AddDimDiametric(){}
    fn AddDimRotated(){}
    fn AddDimOrdinate(){}
    fn AddDimRadial(){}
    fn AddEllipse(){}
    fn AddEllipticalCone(){}
    fn AddEllipticalCylinder(){}
    fn AddExtrudedSolid(){}
    fn AddExtrudedSolidAlongPath(){}
    fn AddLeader(){}
    fn AddMText(){}
    fn AddPoint(){}
    fn AddLightWeightPolyline(){}
    fn AddPolyline(){}
    fn AddRay() {}
    fn AddRegion(){}
    fn AddRevolvedSolid(){}
    fn AddShape(){}
    fn AddSolid(){}
    fn AddSphere(){}
    fn AddSpline(){}
    fn AddText(){}
    fn AddTolerance(){}
    fn AddTorus(){}
    fn AddTrace(){}
    fn AddWedge(){}
    fn AddXline(){}
    fn InsertBlock(){}
    fn AddHatch(){}
    fn AddRaster(){}
    fn AddLine(){}
    fn get_IsLayout(){}
    fn get_Layout(){}
    fn get_IsXRef(){}
    fn AddMInsertBlock(){}
    fn AddPolyfaceMesh(){}
    fn AddMLine(){}
    fn AddDim3PointAngular(){}
    fn get_XRefDatabase(){}
    fn AttachExternalReference(){}
    fn Unload(){}
    fn Reload(){}
    fn Bind(){}
    fn Detach(){}
    fn AddTable(){}
    fn get_Path(){}
    fn put_Path(){}
    fn get_Comments(){}
    fn put_Comments(){}
    fn get_Units(){}
    fn put_Units(){}
    fn get_Explodable(){}
    fn put_Explodable(){}
    fn get_BlockScaling(){}
    fn put_BlockScaling(){}
    fn get_IsDynamicBlock(){}
    fn AddDimArc(){}
    fn AddDimRadialLarge(){}
    fn AddSection(){}
    fn AddMLeader(){}
}