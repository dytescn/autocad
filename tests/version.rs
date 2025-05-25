// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::application;
    use autocad::ffi::version;
    #[test]
    fn test_application_version() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let ver = application::get_version("25".to_string());
        println!("{:?}",ver);
        let vers_list = version::get_version_list();
        println!("{:?}",vers_list);
    }
}