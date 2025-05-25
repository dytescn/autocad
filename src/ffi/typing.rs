use serde_derive::{Deserialize, Serialize};

#[derive(Debug)]
#[derive(Serialize, Deserialize)]
#[derive(Clone)]
pub struct PrintPont{
    pub start_x:f64,
    pub start_y:f64,
    pub end_x:f64,
    pub end_y:f64,
}