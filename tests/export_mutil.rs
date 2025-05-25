// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::export_multi_preview;
    use autocad::ffi::export_preview_file;
    use autocad::ffi::typing::PrintPont;
    #[test]
    fn test_blocks_list() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\";
        let res = export_multi_preview::autocad_export_preview_point("25");
        println!("{:?}",res.unwrap());
    }

    #[test]
    fn test_export_preview_file() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        // start_x: 849.0, start_y: 136.0, end_x: 3383.0, end_y: 2060.0
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\1111";
        let ppt = PrintPont{
            start_x: 849.0,
            start_y: 136.00,
            end_x: 3383.0,
            end_y: 2060.0,
        };
        let res = export_preview_file::autocad_export_preview_point(src_path,1,"25",ppt);
        println!("{:?}",res.unwrap());
    }
}
