// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::draw;
    #[test]
    fn test_autocad_draw() {
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\imgcache\\111111.png";
        let res = draw::autocad_draw("25");
        println!("{:?}",res);
    }
}