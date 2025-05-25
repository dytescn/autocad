use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;

pub struct IAcadUtility {
    disp:ComObject
}

impl IAcadUtility{
    pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    fn AngleToReal(){

    }
    fn AngleToString(){

    }
    fn DistanceToReal(){

    }
    fn RealToString(){

    }
    fn TranslateCoordinates(){

    }
    fn InitializeUserInput(){

    }
    fn GetInteger(){

    }
    fn GetReal(){

    }
    fn GetInput(){

    }
    fn GetKeyword(){

    }
    fn GetString(){

    }
    fn GetAngle(){

    }
    fn AngleFromXAxis(){

    }
    fn GetCorner(){

    }
    fn GetDistance(){

    }
    fn GetOrientation(){

    }
    fn GetPoint(){

    }
    fn PolarPoint(){

    }
    fn CreateTypedArray(){

    }
    fn GetEntity(){

    }
    fn Prompt(){

    }
    fn GetSubEntity(){

    }
    fn IsURL(){

    }
    fn GetRemoteFile(){

    }
    fn PutRemoteFile(){

    }
    fn IsRemoteFile(){

    }
    fn LaunchBrowserDialog(){

    }
    fn SendModelessOperationStart(){

    }
    fn SendModelessOperationEnded(){

    }
    fn GetObjectIdString(){

    }
}