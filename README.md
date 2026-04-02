# PDF 款号提取工具

这是一个用于从PDF检测报告中提取报告编号、款号和测试结果的Python脚本。脚本会处理指定目录下的所有PDF文件，并将提取的数据导出为Excel文件。

## 功能特性

- 从PDF第一页提取报告编号
- 从PDF表格中提取款号（格式：大写字母 + 9位数字，可选-A后缀）
- 提取测试结果
- 支持多个款号，每个款号生成一行Excel记录
- 错误处理和日志输出

## 版本要求

- Python 3.6 或更高版本
- 依赖库：
  - pdfplumber >= 0.5.0
  - pyexcel >= 0.6.0

## 安装步骤

### 1. 初始化虚拟环境

```bash
# 创建虚拟环境
python3 -m venv .venv

# 激活虚拟环境
# Linux/Mac:
source .venv/bin/activate
# Windows:
# .venv\Scripts\activate
```

### 2. 安装依赖

项目使用Poetry管理依赖，请确保已安装Poetry。

```bash
# 安装Poetry（如果尚未安装）
curl -sSL https://install.python-poetry.org | python3 -

# 安装项目依赖
poetry install
```

Poetry会自动创建和管理虚拟环境。如需手动激活虚拟环境：

```bash
poetry shell
```

退出虚拟环境：

```bash
# 如果使用 poetry shell
exit

# 如果手动激活虚拟环境
deactivate
```

## 使用方法

### 运行脚本

```bash
python main.py <input_directory>
```

- `<input_directory>`: 包含PDF文件的目录路径

### 示例

假设PDF文件位于 `input/` 目录下：

```bash
python main.py input
```

脚本会处理 `input/` 目录下的所有 `.pdf` 文件，并生成 `input/output.xlsx` 文件。

### 输出

- 控制台输出：每个文件的处理状态（“文件已处理完毕”或“文件处理失败已跳过”）
- Excel文件：包含报告编号、款号、测试结果的表格，每个款号占一行

### Excel文件格式

| 报告编号        | 款号         | 测试结果  |
|-------------|------------|-------|
| 250345382-2 | M103252250 | 棉 100 |
| 250345382-2 | T104252180 | 棉 100 |
| ...         | ...        | ...   |

## 注意事项

- 确保PDF文件格式正确，包含所需的表格和文本
- 脚本假设“客户认定信息”列包含款号，“测试结果”列包含测试结果
- 如果PDF结构不同，可能需要调整代码中的列名或逻辑
- 处理大量PDF文件时，请确保内存充足

## 许可证

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.
