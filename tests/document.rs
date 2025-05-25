// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use autocad::ffi::document;
    #[test]
    fn test_application_doc_info() {
        let ver = document::document_info("25".to_string());
        println!("{:?}",ver);
    }

    #[test]
    fn test_document_create() {
        // let str_path = "C:\\Users\\Lenovo\\Desktop\\1111test1.dwg";
        let ver = document::document_create();
        println!("{:?}",ver);
    }
    #[test]
    fn test_document_open() {
        let str_path = "C:\\Users\\Lenovo\\Desktop\\2222.dwg";
        let ver = document::document_open(str_path);
        println!("{:?}",ver);
    }

    #[test]
    fn test_document_save() {
        let str_path = "C:\\Users\\Lenovo\\Desktop\\1111test1.dwg";
        let ver = document::document_save(&str_path);
        println!("{:?}",ver);   
    }

    #[test]
    fn test_document_export_cover() {
        let str_path = "C:\\Users\\Lenovo\\Desktop\\";
        let ver = document::document_export_cover(&str_path);
        println!("{:?}",ver);
    }

    #[test]
    fn test_document_info() {
        let res1 = document::get_document_name("25");
        println!("{:?}",res1);
        
        let res2 = document::get_document_path("25");
        println!("{:?}",res2);
        
        let res3 = document::get_document_fullpath("25");
        println!("{:?}",res3);
    }

}