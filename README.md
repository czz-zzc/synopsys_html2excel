# HTML to Excel Register Converter

## 概述

这是一个用于从HTML文件中提取寄存器信息并导出到Excel文件的Python脚本。脚本能够解析包含寄存器定义的HTML文档，提取寄存器的详细信息（如偏移量、位域、访问权限、默认值等），并将这些信息格式化后保存到Excel文件中。

## 功能特性

- 解析HTML文件中的寄存器定义
- 提取寄存器的基本信息（名称、偏移量、大小、描述）
- 提取寄存器字段信息（位域、名称、访问权限、默认值）
- 支持批量处理多个HTML文件
- 支持通配符文件名匹配
- 生成格式化的Excel文件，包含自定义头部信息
- 自动调整Excel列宽
- 将默认值从十六进制格式转换为Verilog格式（如 0xAA → 8'hAA）

## 安装依赖

```bash
pip install beautifulsoup4 pandas openpyxl
```

## 使用方法

### 基本用法

```bash
python reg_html2excel.py input.html
```

### 处理多个文件

```bash
python reg_html2excel.py file1.html file2.html file3.html
```

### 使用通配符

```bash
python reg_html2excel.py *.html
```

### 指定输出文件

```bash
python reg_html2excel.py input.html -o output.xlsx
```

## 命令行参数

- `html_files`: 要解析的HTML文件（支持多个文件和通配符）
- `-o, --output`: 输出Excel文件名（默认：register_output.xlsx）

## 输入文件格式要求

HTML文件应包含以下结构：

1. **寄存器标题**：包含`<strong>`和`<a>`标签的段落
2. **寄存器基本信息**：紧跟在标题后的无序列表，包含：
   - Offset: 寄存器偏移地址
   - Size: 寄存器大小
   - Description: 寄存器描述
3. **字段表格**：包含寄存器字段信息的表格，列包括：
   - 位域范围
   - 字段名称
   - 访问权限
   - 字段描述（包含"Value After Reset"信息）

## 输出格式

生成的Excel文件包含：

### 头部信息（前9行）
- module: 模块名称（从输出文件名提取）
- owner: 所有者（默认：czz）
- size: 大小（默认：64KB）
- cfg_interface: 配置接口（默认：regbus）
- base_addr: 基地址（默认：32'h0000）
- addr_width: 地址宽度（默认：16）
- data_width: 数据宽度（默认：32）
- template_version: 模板版本（默认：v0p0）
- tool_version: 工具版本（默认：v0p0）

### 数据表格（从第10行开始）
| 列名 | 描述 |
|------|------|
| offset | 寄存器偏移地址 |
| reg_name | 寄存器名称 |
| bits | 位域范围 |
| field | 字段名称 |
| sw_access | 软件访问权限（RW/RO） |
| hw_access | 硬件访问权限（通常为空） |
| default | 默认值（Verilog格式） |
| attribute | 属性（默认：normal） |
| description | 描述（通常为空） |

## 数据处理规则

1. **访问权限映射**：
   - "R/W" → "RW"
   - "R" → "RO"
   - 其他 → "unknown"

2. **默认值格式转换**：
   - 0xAA → 8'hAA（其中8是从位域计算得出的宽度）

3. **保留字段处理**：
   - 跳过名称为空或包含"reserved"的字段

## 示例

```bash
# 解析单个HTML文件
python reg_html2excel.py registers.html

# 解析多个HTML文件并指定输出文件名
python reg_html2excel.py reg1.html reg2.html -o my_registers.xlsx

# 使用通配符解析所有HTML文件
python reg_html2excel.py *.html -o all_registers.xlsx
```

## 注意事项

- HTML文件使用ISO-8859-1编码读取
- 输出的Excel工作表名称为"reglist"
- 脚本会自动跳过没有字段名称或标记为"reserved"的字段
- 列宽会根据内容自动调整

## 故障排除

如果脚本无法正确解析HTML文件，请检查：
1. HTML文件的结构是否符合预期格式
2. 寄存器标题是否包含正确的HTML标签
3. 表格是否使用了class="table"的样式
4. 字段信息是否包含"Value After Reset"标