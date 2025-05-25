// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::export_preview;
    #[test]
    fn test_blocks_list() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\";
        let ver = export_preview::autocad_export_preview(src_path.to_string(),"25");
        println!("{:?}",ver);
    }
    #[test]
    fn test_export_sum() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\";
        let ver = export_preview::autocad_export_sum("25");
        println!("{:?}",ver);
    }
    #[test]
    fn test_export_preview() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src_path = "C:\\Users\\Lenovo\\Desktop\\imgcache\\version".to_string();
        let ver = export_preview::autocad_export_preview(src_path,"25");
        println!("{:?}",ver);
    }
}
