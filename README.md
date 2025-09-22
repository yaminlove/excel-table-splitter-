# 📊 表格分割工具

一个基于Streamlit的Web应用，用于将Excel表格按数量总和进行智能分割。

## ✨ 功能特点

- 📊 支持 .xls 和 .xlsx 格式的Excel文件
- 🔄 自动合并连续数量为1的行
- ✂️ 按指定数量总和分割表格
- 📁 生成多个独立的Excel文件并打包下载
- 🎨 现代化的用户界面
- ⚙️ 可自定义数量限制
- 📱 响应式设计，支持手机访问

## 🚀 在线体验

访问 [Streamlit Cloud部署地址](你的部署地址) 直接使用

## 🛠️ 本地运行

### 1. 克隆项目

```bash
git clone https://github.com/你的用户名/表格分割工具.git
cd 表格分割工具
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行应用

```bash
streamlit run streamlit_app.py
```

### 4. 访问网站

打开浏览器访问: http://localhost:8501

## 📋 使用说明

1. **上传文件**: 在网页上选择要处理的Excel文件
2. **设置参数**: 调整数量总和限制（默认590）
3. **处理数据**: 点击"开始处理表格"按钮
4. **预览结果**: 查看分割后的表格信息
5. **下载文件**: 下载生成的ZIP文件

## 🔧 处理逻辑

1. **合并连续数量为1的行**: 连续的数量为1的行会合并为一行，除数量列外其他列置空
2. **按数量分割**: 确保每个分割后的表格数量总和不超过指定限制
3. **文件格式**: 输出为Excel格式的文件，命名为Sheet1.xlsx, Sheet2.xlsx等

## 📁 文件结构

```
├── streamlit_app.py    # Streamlit主应用
├── requirements.txt    # Python依赖
└── README.md          # 说明文档
```

## 📝 文件要求

- Excel文件必须包含"数量"列
- 数量列应为数值型数据
- 支持中文列名和文件名

## 🌐 部署到云端

### Streamlit Cloud (推荐)

1. 将代码推送到GitHub
2. 访问 [share.streamlit.io](https://share.streamlit.io)
3. 连接GitHub仓库
4. 自动部署并获得公网地址

### 其他平台

- Heroku
- Railway
- Render
- Vercel

## 🤝 贡献

欢迎提交Issue和Pull Request！

## 📄 许可证

MIT License