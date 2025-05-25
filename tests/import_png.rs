// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::import_png;
    use autocad::ffi::import_file;
    #[test]
    fn test_autocad_import_png() {
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\11111zaaa_qrcode.png";
        let src_path = "D:\\soft\\piksel\\cache\\dbdc4822c236421280348d5da20af735.png";
        let res = import_png::autocad_import_png(src_path,10000,"25");
        println!("{:?}",res);
    }
    #[test]
    fn test_import_file() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src ="C:\\Users\\Lenovo\\Desktop\\11111.pdf";
        let ver = import_file::autocad_import_file(src,10000,"25");
        println!("{:?}",ver);
    }
}