# Excel转Markdown工具

一个功能强大的在线工具，用于将Excel文件转换为Markdown格式，并支持多文件合并和CSV导出功能。

## 功能特性

### 📊 Excel转Markdown
- 支持单个或多个Excel文件上传
- 智能表格格式转换
- 保持原始数据结构和格式
- 支持中文字符

### 🔄 Excel文件合并
- 多个Excel文件智能合并
- 自动处理表头对齐
- 支持不同结构的表格合并
- 列选择器功能，可自定义合并列

### 📁 CSV导出功能
- 合并Excel文件后直接导出为CSV
- 支持中文字符编码（UTF-8 BOM）
- 智能处理特殊字符（逗号、引号、换行符）
- 动态文件名生成

## 使用方法

### 基本转换
1. 点击"选择Excel文件"按钮上传文件
2. 选择一个或多个Excel文件
3. 点击"转换为Markdown"按钮
4. 系统会自动下载转换后的Markdown文件

### Excel合并
1. 上传至少2个Excel文件
2. 点击"合并Excel文件"按钮
3. 在弹出的列选择器中选择要合并的列
4. 点击"确认合并"完成合并
5. 下载合并后的Excel文件

### CSV导出
1. 上传至少2个Excel文件
2. 点击"合并Excel后转换为CSV"按钮
3. 系统会自动合并文件并导出为CSV格式
4. 下载生成的CSV文件

## 技术特性

- **前端技术**: HTML5, CSS3, JavaScript (ES6+)
- **Excel处理**: SheetJS (XLSX库)
- **文件处理**: 支持.xlsx, .xls格式
- **编码支持**: UTF-8 with BOM，完美支持中文
- **浏览器兼容**: 现代浏览器（Chrome, Firefox, Safari, Edge）

## 本地运行

1. 克隆或下载项目文件
2. 在项目目录中启动HTTP服务器：
   ```bash
   python3 -m http.server 8000
   ```
3. 在浏览器中访问 `http://localhost:8000`

## 文件结构

```
├── index.html          # 主页面
├── script.js           # 核心JavaScript功能
├── style.css           # 样式文件
├── .gitignore          # Git忽略文件
└── README.md           # 项目说明
```

## 更新日志

### v1.1.0
- ✨ 新增：Excel合并后转CSV功能
- 🔧 优化：按钮状态管理和用户提示
- 🐛 修复：数据类型转换问题
- 💡 改进：添加工具提示说明按钮使用条件

### v1.0.0
- 🎉 初始版本发布
- ✨ Excel转Markdown功能
- ✨ 多文件合并功能
- ✨ 列选择器功能

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个项目！