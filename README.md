# PDF转PPT转换器

一个强大的PDF到PowerPoint转换工具，支持高保真图像转换和文本图像分离模式。

## ✨ 功能特性

- 📄 **PDF上传**: 支持拖拽和点击上传PDF文件
- 🖼️ **图像模式**: 高保真转换，完美保持原始布局
- 📝 **分离模式**: 智能分离文本和图像，便于编辑
- ⚙️ **DPI设置**: 可调整图像质量(100-300 DPI)
- 📋 **文本提取**: 独立的文本提取功能
- 💾 **自动下载**: 转换完成后自动下载结果

## 🚀 在线体验

访问: [PDF转PPT转换器](https://your-app-name.onrender.com)

## 🛠️ 本地运行

### 环境要求
- Python 3.8+
- pip

### 安装步骤

1. 克隆项目
```bash
git clone https://github.com/yourusername/pdf2ppt.git
cd pdf2ppt
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 启动服务
```bash
python api.py
```

4. 访问应用
打开浏览器访问: http://localhost:8000

## 📁 项目结构

```
pdf2ppt/
├── api.py              # FastAPI后端服务
├── converter.py        # PDF转换核心逻辑
├── requirements.txt    # Python依赖
├── render.yaml        # Render部署配置
├── web/               # 前端文件
│   ├── index.html     # 主页面
│   ├── script.js      # 交互逻辑
│   └── style.css      # 样式文件
└── README.md          # 项目说明
```

## 🔧 API接口

### 转换PDF为PPT
```
POST /convert
Content-Type: multipart/form-data

参数:
- file: PDF文件
- mode: 转换模式 (image/separated)
- dpi: 图像质量 (100-300)
```

### 提取PDF文本
```
POST /extract_text
Content-Type: multipart/form-data

参数:
- file: PDF文件
```

## 📦 技术栈

- **后端**: FastAPI + Python
- **PDF处理**: PyMuPDF
- **PPT生成**: python-pptx
- **图像处理**: Pillow
- **前端**: HTML5 + CSS3 + JavaScript

## 🌟 部署

### Render部署 (推荐)
1. Fork此项目到您的GitHub
2. 在Render创建新的Web Service
3. 连接GitHub仓库
4. 选择免费套餐
5. 自动部署完成

### 其他平台
- Railway
- Vercel (需要调整为Serverless函数)
- Heroku

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📞 联系

如有问题，请提交Issue或联系开发者。