// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::comps_cover;
    #[test]
    fn test_comps_cover() {
        let file_src: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\123.jpg";
        let ver = comps_cover::export_comps_cover(file_src,"25");
        println!("{:?}",ver);
    }
}