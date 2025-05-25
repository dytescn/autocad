// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::import_dwg;
    #[test]
    fn test_autocad_import_dwg() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\222222.dwf";
        let res = import_dwg::autocad_import_dwg(src_path,"25");
        println!("{:?}",res);
    }
}