# Excel操作工具

一个Python编写的Excel操作工具，包含以下功能：

1. **合并多个Excel文件**：支持处理每个Excel文件中的多个sheet页
2. **按列拆分Excel文件**：根据指定的一列或多列进行拆分

## 依赖安装

本工具使用`openpyxl`库来处理Excel文件，需要先安装：

```bash
pip install openpyxl
```

## 使用方法

### 1. 合并Excel文件

```bash
python excel_tool.py merge <输出文件> <输入文件1> <输入文件2> ...
```

**示例**：
```bash
python excel_tool.py merge merged_result.xlsx file1.xlsx file2.xlsx file3.xlsx
```

### 2. 按列拆分Excel文件

```bash
python excel_tool.py split <输入文件> <输出目录> <列索引1> <列索引2> ...
```

**示例**：
```bash
# 按第1列（索引为0）拆分
python excel_tool.py split data.xlsx output_dir 0

# 按第1列和第2列（索引为0和1）拆分
python excel_tool.py split data.xlsx output_dir 0 1
```

## 功能说明

### 合并功能
- 支持合并多个Excel文件
- 自动处理每个文件中的多个sheet页
- 保持原有的表头结构
- 支持处理不同文件中相同名称的sheet页

### 拆分功能
- 支持按一列或多列进行拆分
- 自动创建输出目录
- 为每个拆分结果生成独立的Excel文件
- 文件名包含sheet名称和拆分列的值

## 注意事项

- 仅支持`.xlsx`格式的Excel文件
- 列索引从0开始计数
- 拆分时，列值会被转换为字符串用于生成文件名
- 大文件可能会占用较多内存
