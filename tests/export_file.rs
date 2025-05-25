// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::export_file;
    #[test]
    fn test_application_quit() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src ="D:\\codes\\dytes\\autocad\\cache\\888.dwg";
        let ver = export_file::autocad_export_file(src.to_string(),"25");
        println!("{:?}",ver);
    }
    #[test]
    fn test_export_dxf() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src ="C:\\Users\\Lenovo\\Desktop\\imgcache\\111222.dxf";
        let ver = export_file::autocad_export_dxf(src.to_string(),"25");
        println!("{:?}",ver);
    }
}