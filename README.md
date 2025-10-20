# DOCX to JSON Converter

一个智能的DOCX文档转JSON工具，专门用于提取和组织文档中的图像、文本内容，并实现高级的图像分组功能。

## 特性

### 🎯 核心功能
- **智能图像分组**：自动识别并行（row）和垂直相邻（column）的图像组合
- **真实资源提取**：从DOCX文件中提取真实图像并生成SHA256哈希ID
- **完整内容提取**：支持段落文本和表格内容的完整提取
- **标题与来源归属**：智能分配图像标题和来源信息
- **文档顺序保持**：按原文档顺序输出所有内容块
- **繁体转简体**：支持繁体中文自动转换为简体中文（保护英文、数字、URL等）
- **显式标记系统**：支持"标题"/"Title"和"来源"/"Source"显式标记，提高检测准确性
- **统一来源格式**：所有来源信息统一添加"Source: "前缀

### 🔍 显式标记系统 (新特性)

项目现已支持显式标记系统，大幅提升标题和来源检测的准确性：

#### 显式标记格式
- **标题标记**：段落开头包含"标题"或"Title"
- **来源标记**：段落开头包含"来源"或"Source"

#### 检测优先级
1. **显式标记优先**：优先检测带有标记的段落
2. **启发式fallback**：如无显式标记，使用优化后的启发式算法
3. **智能内容提取**：自动去除标记前缀，提取真实内容

#### 使用建议
在DOCX文档中直接添加标记，例如：
```
标题：市场情绪分析图表
来源：Bloomberg, Financial Times
```

### 🧠 二阶段分组算法

#### Phase 1: 同段落分组 (Row Layout)
- 同一段落内的多张图像自动组合为`layout='row'`
- 典型场景：并排对比图表

#### Phase 2: 相邻段落分组 (Column Layout)  
- 连续段落中的图像根据间隔和文本量进行分组
- 仅当间隔≤`max_gap_paras`且无大量文本(>max_title_len chars)时分组
- 特殊规则：若组合宽度≤`page_width_ratio` * 页面宽度，则使用`layout='row'`

#### 标题与来源归属（增强版）
- **显式标记优先**：检测"标题"/"Title"和"来源"/"Source"标记的段落
- **智能启发式fallback**：标题搜索图片前方文本(≤max_title_len chars)，来源搜索图片后方文本
- **统一格式**：所有来源统一添加"Source: "前缀

## 安装

### 环境要求
- Python 3.9+（推荐 3.12 或 3.13）
- python-docx 库
- opencc-python-reimplemented 库（繁体转简体功能）
- lxml 库

### 安装步骤

#### 快速安装（推荐）
```bash
# 1. 克隆项目
git clone <repository-url>
cd docx2json

# 2. 创建虚拟环境
python3 -m venv .venv

# 3. 激活虚拟环境
source .venv/bin/activate  # macOS/Linux
# Windows: .venv\Scripts\activate

# 4. 安装依赖
pip install -r requirements.txt
```

#### 手动安装
```bash
# 如果没有 requirements.txt，可以手动安装
pip install python-docx opencc-python-reimplemented lxml
```

### 验证安装
```bash
# 确保虚拟环境已激活
python -c "from docx import Document; import opencc; print('✓ 依赖安装成功！')"
```

## 使用方法

### 基本用法
```bash
python to_ncj.py "document.docx"
```

### 完整参数
```bash
python to_ncj.py "input.docx" [options]
  --out content.json              # 输出JSON文件 (默认: content.json)
  --assets-dir assets/media       # 图像资源目录 (默认: assets/media)
  --max_title_len 60              # 标题检测最大字符数 (默认: 60)
  --max_gap_paras 1               # 分组最大段落间隔 (默认: 1) 
  --page_width_ratio 0.95         # 行布局检测宽度比例 (默认: 0.95)
  --traditional-to-simplified     # 启用繁体转简体 (保护英文/数字/URL)
  --no-explicit-markers           # 禁用显式标记检测，仅使用启发式算法
  --debug                         # 在输出中包含分组推理信息
```

### 示例
```bash
# 基本转换
python to_ncj.py "报告.docx" --out report.json

# 带繁体转简体的转换
python to_ncj.py "繁體文檔.docx" --traditional-to-simplified --assets-dir assets/

# 带调试信息的转换
python to_ncj.py "分析.docx" --debug --assets-dir images/

# 仅使用启发式算法（禁用显式标记）
python to_ncj.py "文档.docx" --no-explicit-markers --max_gap_paras 2
```

## 输出格式

### ⚠️ 文件命名规范

**重要提示**：生成的 JSON 文件名**不能包含特殊字符**（如双引号 `"`、问号 `?`、冒号 `:` 等），否则会导致下游脚本读取失败。

**推荐命名格式**：
```bash
# ✅ 正确示例
251020-何为上涨十月_zh-CN.json
251020-what-uptober_en-US.json
251020-何為上漲十月_zh-HK.json

# ❌ 错误示例（包含特殊字符）
251020-何为"上涨十月"？_zh-CN.json  # 双引号和问号会导致 Shell 参数解析失败
report:analysis_zh-CN.json            # 冒号在某些系统中不被允许
```

**命名建议**：
- 使用日期前缀：`YYMMDD-` 或 `YYYYMMDD-`
- 使用下划线或连字符分隔：`-`、`_`
- 语言标识符后缀：`_zh-CN`、`_en-US`、`_zh-HK`
- 避免使用：`"` `?` `:` `*` `<` `>` `|` `/` `\`

### JSON结构
```json
{
  "doc": {
    "title": "文档标题",
    "date": "2025-08-18", 
    "locale": "zh-CN",
    "version": "v1",
    "source_file": "original.docx"
  },
  "blocks": [
    {
      "type": "paragraph",
      "text": "段落文本内容"
    },
    {
      "type": "figure",
      "image": {"asset_id": "img_76f7bfb095b6"},
      "title": "图表标题",
      "credit": "数据来源",
      "group_id": "grp_0001",
      "group_seq": 1,
      "group_len": 2,
      "layout": "row"
    }
  ],
  "assets": [
    {
      "asset_id": "img_76f7bfb095b6",
      "filename": "assets/img_76f7bfb095b6.png",
      "sha256": "76f7bfb095b6f67f8cc5c56857be9ef285cf4065..."
    }
  ],
  "report": {
    "warnings": [],
    "debug": ["grp_0001: row by same-paragraph(para=4, 2 images)..."]
  }
}
```

### 字段说明
- **group_id**: 唯一分组标识符
- **group_seq**: 在组内的序号(1开始)  
- **group_len**: 组内图像总数
- **layout**: 布局类型(`"row"`并行 | `"column"`垂直)
- **asset_id**: 基于SHA256的真实图像ID

## 技术架构

### 核心模块
1. **图像提取** (`extract_figures_from_docx`): 从段落和表格中提取图像
2. **资源处理** (`extract_and_hash_images`): ZIP解压和SHA256哈希计算  
3. **智能分组** (`group_figures`): 二阶段分组算法实现
4. **属性分配** (`assign_titles_and_credits`): 标题和来源的智能归属

### 设计特点
- **分离式block设计**：保持文档线性结构，便于顺序渲染
- **可配置参数**：灵活调整分组行为以适应不同文档类型
- **真实资源管理**：避免placeholder，确保资源完整性
- **调试友好**：提供详细的分组推理信息

## 开发

### 项目结构
```
docx2json/
├── to_ncj.py              # 主转换脚本
├── README.md              # 项目文档
├── assets/                # 提取的图像资源
│   ├── 250818_summer_break/
│   └── 250804_negative_revisions/
├── *.docx                 # 测试文档
├── *.json                 # 转换结果
└── venv/                  # 虚拟环境
```

### 测试
项目包含两个测试用例：
- `250818 - Summer Break.docx`: 14图像，11分组，3个多图分组
- `250804 - Negative Revisions.docx`: 25图像，15分组，10个多图分组

### 扩展
算法支持以下扩展：
- 新的布局类型检测
- 自定义标题/来源匹配模式  
- 多语言文档支持
- 更复杂的图像排列识别

## 常见问题

### Q: 为什么有些图像没有标题？
A: 标题分配基于附近文本的长度和位置。建议使用显式标记（"标题："或"Title:"）来提高检测准确性。

### Q: 如何调整分组敏感度？
A: 使用`--max_gap_paras`调整段落间隔容忍度，使用`--max_title_len`调整标题检测长度。

### Q: 输出的图像文件在哪里？
A: 图像保存在`--assets-dir`指定的目录中，文件名使用SHA256哈希确保唯一性。每个文档的图像将保存在以源文件名命名的子目录中。

### Q: 繁体转简体功能是否会影响英文内容？
A: 不会。该功能使用智能保护机制，仅转换中文字符，保护英文文本、数字、URL和标点符号不被修改。

### Q: 显式标记系统是否兼容旧文档？
A: 完全兼容。如果文档中没有显式标记，系统会自动使用优化后的启发式算法作为fallback。

### Q: 所有来源信息都会添加"Source: "前缀吗？
A: 是的。无论原文档中来源格式如何，输出时都会统一添加"Source: "前缀，确保格式一致性。

## License

MIT License - 详见 LICENSE 文件

## 贡献

欢迎提交 Issues 和 Pull Requests！

---

*该项目专为需要精确图像分组和内容提取的文档处理场景设计*