# xyx-tool

一个基于 Spring Boot 的 Excel 报表处理工具，用于上传屏风业务工作簿并生成日报或月报。处理完成后，服务会将补全计算列的 `.xlsx` 文件作为下载响应返回。

## 功能

- 提供浏览器上传页面，可分别生成屏风日报和月报。
- 读取 Excel 工作簿中的业务数据、订单和物流信息。
- 自动补充店铺、成本、物流、客服提成、推广、总成本、金额和扇数等报表列。
- 支持合并单元格场景。
- 支持最大 100 MB 的单个上传文件及请求体。

## 技术栈

- Java 8+
- Spring Boot 2.2.5
- Maven
- Thymeleaf
- Apache POI、EasyExcel、Hutool

## 快速开始

### 前置条件

- JDK 8 或更高版本
- Maven 3.6 或更高版本

### 构建与启动

```bash
mvn clean package
mvn spring-boot:run
```

服务默认运行在 `http://localhost:9091`。

启动后访问：

```text
http://localhost:9091/shop/gotoUploadPage
```

在页面中选择 Excel 文件，并按需要提交“日报”或“月报”任务；浏览器会下载生成的结果文件。

## 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/shop/test` | 健康检查，返回 `hello test content`。 |
| GET | `/shop/gotoUploadPage` | 打开文件上传页面。 |
| POST | `/shop/generatorScreenDayReport` | 上传字段 `file`，生成屏风日报并下载。 |
| POST | `/shop/generatorScreenMonthReport` | 上传字段 `file`，生成屏风月报并下载。 |

也可以使用 `curl` 调用日报接口：

```bash
curl -X POST http://localhost:9091/shop/generatorScreenDayReport \
  -F "file=@input.xlsx" \
  -o screen-day-report.xlsx
```

## Excel 模板要求

该工具按固定模板的工作表名称和列布局读取数据。上传前请确认工作簿符合对应任务的模板。

- 日报：主数据工作表必须命名为 `Sheet1`，订单/物流关联数据使用 `Sheet2`。
- 月报：主数据工作表必须命名为 `总表`。
- 输入文件应为 `.xlsx`；生成结果会保留原始文件名并附加时间戳。

缺少必需工作表、列布局不匹配或单元格数据无法解析时，服务会返回错误页面或请求异常信息。

## 配置

默认配置位于 [application.yaml](D:/project/xyx-tool/src/main/resources/application.yaml)：

```yaml
server:
  port: 9091

spring:
  servlet:
    multipart:
      max-file-size: 100MB
      max-request-size: 100MB
```

如需修改端口或上传大小限制，请调整该文件后重启服务。应用还将嵌入式 Tomcat 的 `maxSwallowSize` 设置为无限制，以避免大文件上传时连接被重置。

## 测试

```bash
mvn test
```

当前测试为基础 Maven/JUnit 测试；实际报表规则建议使用脱敏的模板样例进行人工校验。

## 项目结构

```text
src/main/java/org/example/
├── Application.java              # Spring Boot 启动类与 Tomcat 上传配置
├── controller/ShopController.java # 页面与报表生成接口
├── service/ShopService.java       # Excel 读取、计算与导出逻辑
├── dto/                           # 报表数据传输对象
└── handler/MyExceptionHandler.java # 全局异常处理

src/main/resources/
├── application.yaml               # 服务与上传配置
└── templates/                     # Thymeleaf 页面模板
```

## 注意事项

- 报表计算逻辑依赖业务模板中固定的列索引，请避免任意调整列顺序。
- 请勿上传包含敏感业务数据的文件到不受信任的部署环境。
