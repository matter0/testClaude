package org.jeecgframework.boot3.importexceltodb;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ImportExcelToDbApplication {

    // Excel文件所在文件夹路径
    private static final String EXCEL_FOLDER_PATH = "d:\\Users\\31234\\Desktop\\dataSource";

    public static void main(String[] args) {
        // 启动Spring Boot应用
        SpringApplication.run(ImportExcelToDbApplication.class, args);

        // 创建Excel导入服务实例并执行导入
        ExcelImportService importService = new ExcelImportService();
        try {
            importService.importExcelToDb(EXCEL_FOLDER_PATH);
        } catch (Exception e) {
            System.err.println("数据导入失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

}
