// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::plotcfg;
    #[test]
    fn test_plotcfg_change() {
        // let src_path: &str = "C:\\Users\\Lenovo\\Desktop\\";
        let ver = plotcfg::plot_cfg("25");
        println!("{:?}",ver);
    }
}