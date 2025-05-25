use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

use super::block::IAcadBlocks;
use super::group::IAcadGroups;
use super::layer::IAcadLayers;
use super::layouts::IAcadLayouts;
use super::plotconfigurations::IAcadPlotConfigurations;
use super::plotconfigurations;
use super::modelspace::IAcadModelSpace;

pub struct IAcadDatabase {
    disp:ComObject
}

impl IAcadDatabase{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    pub fn get_modelspace(&self)->Option<IAcadModelSpace>{
        let doc_res = self.disp.get_property("modelspace");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let info = IAcadModelSpace::new(disp.clone());
                        return Some(info);
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
    fn get_PaperSpace(){

    }
    pub fn get_Blocks(&self)->Option<IAcadBlocks>{
        let doc_res = self.disp.get_property("blocks");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let info = IAcadBlocks::new(disp.clone());
                        return Some(info);
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
    fn CopyObjects(){

    }
    pub fn get_Groups(&self)->Option<IAcadGroups>{
        let doc_res = self.disp.get_property("groups");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let info = IAcadGroups::new(disp.clone());
                        return Some(info);
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
    fn get_DimStyles(){

    }
    pub fn get_Layers(&self)->Option<IAcadLayers>{
        let info_res = self.disp.get_property("layers");
        match info_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let info = IAcadLayers::new(disp.clone());
                        return Some(info);
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
    fn get_Linetypes(){

    }
    fn get_Dictionaries(){

    }
    fn get_RegisteredApplications(){

    }
    fn get_TextStyles(){

    }
    fn get_UserCoordinateSystems(){

    }
    fn get_Views(){

    }
    fn get_Viewports(){

    }
    fn get_ElevationModelSpace(){

    }
    fn put_ElevationModelSpace(){

    }
    fn get_ElevationPaperSpace(){

    }
    fn put_ElevationPaperSpace(){

    }
    fn get_Limits(){

    }
    fn put_Limits(){

    }
    fn HandleToObject(){

    }
    fn ObjectIdToObject(){

    }
    pub fn get_Layouts(&self)->Option<IAcadLayouts>{
        let doc_res = self.disp.get_property("layouts");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let acad_doc = IAcadLayouts::new(disp.clone());
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
    pub fn get_PlotConfigurations(&self)->Option<IAcadPlotConfigurations>{
        let hres = self.disp.get_property("plotconfigurations").log_error("got err").unwrap();
        match hres.to_idispatch(){
            Ok(disp)=>{
                let acad_doc = IAcadPlotConfigurations::new(disp.clone());
                return Some(acad_doc);
            }
            Err(e)=>{
                return None
            }
        }
    }
    fn get_Preferences(){

    }
    fn get_SummaryInfo(){

    }
    fn get_SectionManager(){

    }
    fn get_Materials(){

    }
}