// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::application;
    #[test]
    fn test_application_quit() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let ver = application::do_quit("25".to_string());
        println!("{:?}",ver);
    }
    #[test]
    fn test_application_info() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let ver = application::get_cad_execute_path("25".to_string());
        println!("{:?}",ver);
    }
    #[test]
    fn test_application_open() {
        let src ="C:\\Users\\Lenovo\\Desktop\\1111.dwg".to_string();
        let ver = application::do_open(src,"25".to_string());
        println!("{:?}",ver);
    }
}