package org.jeecgframework.boot3.importexceltodb;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel数据导入服务类
 * 负责将Excel文件数据导入到MySQL数据库
 */
public class ExcelImportService {

    // 数据库配置 (请替换为你自己的)
    private static final String DB_URL = "jdbc:mysql://localhost:3306/testimport?useSSL=false&characterEncoding=utf8";
    private static final String DB_USER = "root";
    private static final String DB_PASS = "123456";

    // 数据库插入SQL
    private static final String INSERT_SQL =
            "INSERT INTO dm_factory_record (train_type_code, train_type, train_code, coach_no, material_name, serial_number, manufacturer_name) VALUES (?, ?, ?, ?, ?, ?, ?)";

    /**
     * 执行Excel数据导入
     * @param excelFolderPath Excel文件所在的文件夹路径
     * @return 导入的总记录数
     */
    public int importExcelToDb(String excelFolderPath) {
        // 存放Excel文件的文件夹路径
        File folder = new File(excelFolderPath);
        File[] listOfFiles = folder.listFiles((dir, name) -> name.endsWith(".xlsx"));

        if (listOfFiles == null || listOfFiles.length == 0) {
            System.out.println("未找到 Excel 文件！");
            return 0;
        }

        try (Connection conn = DriverManager.getConnection(DB_URL, DB_USER, DB_PASS)) {
            conn.setAutoCommit(false); // 开启事务
            try (PreparedStatement pstmt = conn.prepareStatement(INSERT_SQL)) {

                int totalInserted = 0;
                for (File file : listOfFiles) {
                    System.out.println("正在处理文件: " + file.getName());
                    totalInserted += processSingleFile(file, pstmt);
                }

                conn.commit(); // 提交全部事务
                System.out.println("导入完成！共插入 " + totalInserted + " 条记录。");
                return totalInserted;

            } catch (Exception e) {
                conn.rollback();
                e.printStackTrace();
                throw e;
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("导入数据失败", e);
        }
    }

    /**
     * 处理单个Excel文件
     * @param file Excel文件
     * @param pstmt 预处理语句
     * @return 本文件插入的记录数
     * @throws Exception 处理异常
     */
    private int processSingleFile(File file, PreparedStatement pstmt) throws Exception {
        int batchCount = 0;
        String fileName = file.getName();

        // 1. 解析文件名中的规则 (车型编码 和 车型)
        String cxbm = "";
        String cx = "";
        if (fileName.contains("12号线")) {
            cxbm = "DL12";
            cx = "大连12号线";
        } else if (fileName.contains("1号线版") || fileName.contains("1号线")) {
            cxbm = "DL01";
            cx = "大连1号线";
        } else if (fileName.contains("2号线版") || fileName.contains("2号线")) {
            cxbm = "DL02";
            cx = "大连2号线";
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                String sheetName = sheet.getSheetName(); // 例: "1车"

                // 2. 解析车厢号 (如 "1车" -> "01")
                String cxh = "";
                Matcher m = Pattern.compile("(\\d+)").matcher(sheetName);
                if (m.find()) {
                    cxh = String.format("%02d", Integer.parseInt(m.group(1)));
                }

                // 3. 寻找并解析 "车号" (通常在前5行内)
                String ch = "";
                for (int i = 0; i < 5; i++) {
                    Row r = sheet.getRow(i);
                    if (r != null && r.getCell(0) != null) {
                        String cellVal = getCellValue(r.getCell(0));
                        if (cellVal.contains("车号：")) {
                            // 截取 "车号：0129-1" 中的 "0129"
                            String[] parts = cellVal.replace("车号：", "").split("-");
                            if (parts.length > 0) ch = parts[0].trim();
                            break;
                        }
                    }
                }

                // 4. 核心解析：遍历数据行 (主从合并拆分逻辑 + 合并单元格厂家继承)
                StringBuilder currentName = new StringBuilder();
                List<String> currentSerials = new ArrayList<>();
                String blockFactory = ""; // 当前正在收集的部件块的厂家
                String globalLatestFactory = ""; // 【新增】记忆最新读取到的有效厂家名称（解决Excel合并单元格问题）

                // 从大概第4或第5行开始遍历（跳过表头）
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) continue;

                    String col0_xuhao = getCellValue(row.getCell(0)); // 序号
                    String col1_mingcheng = getCellValue(row.getCell(1)); // 名称
                    String col2_xuliehao = getCellValue(row.getCell(2)); // 序列号
                    String col3_changjia = getCellValue(row.getCell(3)); // 制造工厂

                    // 判断是否是结构行（如 "车下 ：", "车上 ：", 或是表头 "序号"）
                    boolean isStructuralRow = col0_xuhao.contains("车下") || col0_xuhao.contains("车上") || col0_xuhao.contains("序号");

                    // 【触发器】：当遇到新的序号数字，或者遇到结构分隔行时，结算上一个部件块
                    if (col0_xuhao.matches("\\d+") || isStructuralRow) {
                        // 结算当前缓存的块，存入批处理（注意这里传入的是 blockFactory）
                        batchCount += flushBlockToDb(pstmt, cxbm, cx, ch, cxh, currentName.toString(), currentSerials, blockFactory);

                        // 清空缓存，准备迎接新部件
                        currentName.setLength(0);
                        currentSerials.clear();
                        blockFactory = ""; // 结算后重置当前块的厂家
                    }

                    // 如果是结构行，跳过当前行数据累加
                    if (isStructuralRow || col0_xuhao.contains("车号")) {
                        // 【安全措施】：遇到"车上""车下"等大区域分割时，清空记忆的厂家，防止车上的件错误继承车下的厂家
                        globalLatestFactory = "";
                        continue;
                    }

                    // ====== 【核心修改：处理合并单元格 / 厂家继承逻辑】 ======
                    // 1. 如果当前行读取到了真实的厂家内容，就更新"记忆体"
                    if (!col3_changjia.isEmpty()) {
                        globalLatestFactory = col3_changjia;
                    }
                    // 2. 如果当前块还没有厂家，就自动继承"记忆体"里的最新厂家
                    if (blockFactory.isEmpty() && !globalLatestFactory.isEmpty()) {
                        blockFactory = globalLatestFactory;
                    }
                    // =========================================================

                    // 【累加器】：将当前行的名称和序列号累加到块缓存中
                    if (!col1_mingcheng.isEmpty()) currentName.append(col1_mingcheng); // 名称拼接
                    if (!col2_xuliehao.isEmpty()) currentSerials.add(col2_xuliehao); // 序列号加入集合
                }

                // 遍历完整个 Sheet 后，千万别忘了结算最后一个部件块
                batchCount += flushBlockToDb(pstmt, cxbm, cx, ch, cxh, currentName.toString(), currentSerials, blockFactory);
            }

            // 执行批量插入 (每个文件执行一次，防止内存溢出)
            pstmt.executeBatch();
        }
        return batchCount;
    }

    /**
     * 结算块数据并加入到 JDBC 批处理中
     * @param pstmt 预处理语句
     * @param cxbm 车型编码
     * @param cx 车型
     * @param ch 车号
     * @param cxh 车厢号
     * @param wlmc 物料名称
     * @param serials 序列号列表
     * @param zzsmc 制造商名称
     * @return 插入的记录数
     * @throws Exception 数据库异常
     */
    private int flushBlockToDb(PreparedStatement pstmt, String cxbm, String cx, String ch, String cxh,
                              String wlmc, List<String> serials, String zzsmc) throws Exception {

        // 【修改点1】：拦截条件改变。不再以"有没有序列号"为准，而是看"有没有部件名称"。
        // 如果连部件名称都没有，说明是纯空行，直接跳过不入库。
        if (wlmc == null || wlmc.trim().isEmpty()) {
            return 0;
        }

        int count = 0;

        // 【修改点2】：如果没有序列号，依然正常插入一条记录，序列号留空
        if (serials.isEmpty()) {
            pstmt.setString(1, cxbm);
            pstmt.setString(2, cx);
            pstmt.setString(3, ch);
            pstmt.setString(4, cxh);
            pstmt.setString(5, wlmc);
            pstmt.setString(6, ""); // 序列号写入空字符串
            pstmt.setString(7, zzsmc);
            pstmt.addBatch();
            count++;
        } else {
            // 如果有序列号，规则不变：有几个序列号，就拆分成几条记录
            for (String xlh : serials) {
                pstmt.setString(1, cxbm);
                pstmt.setString(2, cx);
                pstmt.setString(3, ch);
                pstmt.setString(4, cxh);
                pstmt.setString(5, wlmc);
                pstmt.setString(6, xlh); // 填入对应的序列号
                pstmt.setString(7, zzsmc);
                pstmt.addBatch();
                count++;
            }
        }

        return count;
    }

    /**
     * 安全获取单元格字符串值的工具方法
     * @param cell Excel单元格
     * @return 单元格字符串值
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        cell.setCellType(CellType.STRING); // 强制转换为字符串读取，避免数字格式导致的小数点问题
        return cell.getStringCellValue().trim();
    }
}