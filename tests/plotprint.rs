#[cfg(test)]
mod tests {
    use autocad::ffi::plotprint;
    #[test]
    fn test_application_doc_info() {
        let src_path = "C:\\Users\\Lenovo\\Desktop\\imgcache\\123";
        let ver = plotprint::batch_print_file(src_path,"25");
        println!("{:?}",ver);
        // for i in 0..10 {
        //     sleep(Duration::from_secs(1));
        //     println!("{:?}",i);
        // }
    }
}
