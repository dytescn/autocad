// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::layer;
    #[test]
    fn test_layout_preview() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\";
        let ver = layer::layer_plot_preview(src_path);
        println!("{:?}",ver);
    }
}