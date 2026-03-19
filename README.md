# Excel数据导入数据库工具

## 项目简介
这是一个基于Spring Boot的Java应用程序，用于将大连地铁的Excel数据文件批量导入到MySQL数据库中。

## 主要功能
1. **批量处理Excel文件**：自动扫描指定目录下的所有.xlsx文件
2. **智能解析**：根据文件名自动识别线路（1号线、2号线、12号线）
3. **多Sheet处理**：支持Excel文件中的多个工作表（如"1车"、"2车"等）
4. **数据清洗**：自动处理合并单元格、空值、数据格式等问题
5. **事务管理**：支持批量插入和事务回滚
6. **错误处理**：完善的异常处理和日志记录

## 支持的数据格式
- **Excel文件命名**：包含"1号线"、"2号线"、"12号线"等标识
- **工作表命名**：如"1车"、"2车"等，自动提取车厢号
- **数据列**：
  - 序号
  - 名称（部件名称）
  - 序列号
  - 制造工厂

## 数据库表结构
数据将导入到`dm_factory_record`表，包含以下字段：
- `train_type_code` - 车型编码（如DL01、DL02、DL12）
- `train_type` - 车型（如大连1号线、大连2号线、大连12号线）
- `train_code` - 车号（从Excel中提取）
- `coach_no` - 车厢号（从工作表名提取，如01、02）
- `material_name` - 物料名称（部件名称）
- `serial_number` - 序列号（可为空）
- `manufacturer_name` - 制造商名称

## 使用方法
### 环境要求
- Java 17+
- MySQL 8.0+
- Maven 3.6+

### 配置步骤
1. 创建数据库表：
```sql
CREATE TABLE dm_factory_record (
    id INT AUTO_INCREMENT PRIMARY KEY,
    train_type_code VARCHAR(20),
    train_type VARCHAR(50),
    train_code VARCHAR(20),
    coach_no VARCHAR(10),
    material_name VARCHAR(500),
    serial_number VARCHAR(100),
    manufacturer_name VARCHAR(200),
    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

2. 修改数据库配置（在`ImportExcelToDbApplication.java`中）：
```java
private static final String DB_URL = "jdbc:mysql://localhost:3306/your_database";
private static final String DB_USER = "your_username";
private static final String DB_PASS = "your_password";
```

3. 准备Excel文件：
   - 将Excel文件放入`d:\Users\31234\Desktop\dataSource\`目录
   - 或修改源代码中的文件路径

### 运行程序
```bash
# 编译打包
mvn clean package

# 运行程序
java -jar target/ImportExcelToDb-0.0.1-SNAPSHOT.jar
```

## 项目结构
```
ImportExcelToDb/
├── src/main/java/org/jeecgframework/boot3/importexceltodb/
│   └── ImportExcelToDbApplication.java  # 主程序
├── src/main/resources/
│   └── application.properties           # 配置文件
├── pom.xml                              # Maven依赖配置
└── README.md                            # 项目说明
```

## 技术栈
- **Spring Boot 4.0.3** - 应用框架
- **Apache POI 5.2.3** - Excel文件处理
- **MySQL Connector 8.0.33** - 数据库连接
- **Maven** - 项目构建

## 注意事项
1. Excel文件必须使用.xlsx格式
2. 数据库连接信息需要根据实际环境修改
3. 程序运行时需要确保数据库服务已启动
4. 建议在测试环境中先验证数据导入结果

## 问题反馈
如遇到问题，请检查：
1. 数据库连接配置是否正确
2. Excel文件格式是否符合要求
3. 文件路径是否存在
4. 数据库表结构是否创建正确

## 版本信息
- 当前版本：0.0.1-SNAPSHOT
- 更新时间：2026年3月19日