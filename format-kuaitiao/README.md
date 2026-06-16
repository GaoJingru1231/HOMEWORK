# 格式快调 - 论文自动排版网站

基于「新闻3班 Triple F 实践一选题」构建的文档格式一键标准化工具。

## 项目结构

```
output/
├── index.html        # 前端页面（单文件，内嵌 CSS/JS）
├── server.py         # 后端 API 服务（FastAPI + python-docx）
└── requirements.txt  # Python 依赖
```

## 快速启动

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动后端服务

```bash
cd output
python server.py
```

服务默认运行在 http://localhost:8000

### 3. 打开前端

浏览器打开 `output/index.html`，或直接访问 http://localhost:8000

## 功能说明

| 功能 | 说明 |
|------|------|
| 文件上传 | 支持 .docx / .doc 拖拽上传，最大 50MB |
| AI 格式检测 | 自动检测字体、行距、缩进、页边距、标题层级等 6 大类问题 |
| 内置模板 | 12 套校园高频场景预设模板（毕业论文、课程论文、实验报告等） |
| 自定义格式 | 20+ 格式参数独立调节，支持保存为个人模板 |
| 一键排版 | 按模板全文批量修改格式，生成新文档下载 |
| 个人模板 | 本地持久化存储自定义模板，支持编辑和删除 |

## 技术栈

- **前端**：HTML5 + CSS3 + Vanilla JS（单文件架构）
- **后端**：Python FastAPI + python-docx
- **文档处理**：python-docx（段落/字体/页边距/标题层级修改）
- **排版引擎**：规则引擎 + 模板映射，批量修改全文格式

## API 接口

| 端点 | 方法 | 说明 |
|------|------|------|
| `/api/health` | GET | 健康检查 |
| `/api/templates` | GET | 获取预设模板列表 |
| `/api/detect` | POST | 上传文档 → 格式检测报告 |
| `/api/format` | POST | 上传文档 → 一键排版 → 下载链接 |
| `/api/download/{id}/{file}` | GET | 下载排版后文档 |

## 选题信息

- **小组**：Triple F（吴明煊、高静茹、张艳玉）
- **选题**：「格式快调」——高校师生通用文档格式一键标准化工具网站
- **班级**：新闻3班 实践一
