// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::layout;
    #[test]
    fn test_layout_preview() {
        let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\123123.jpg";
        let ver = layout::layout_plot_preview(src_path);
        println!("{:?}",ver);
    }
}
