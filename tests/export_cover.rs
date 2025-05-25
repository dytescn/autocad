// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::export_cover;
    #[test]
    fn test_blocks_list() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\123123.jpg";
        let ver = export_cover::autocad_export_cover(src_path.to_string(),"25");
        println!("{:?}",ver);
    }
}

