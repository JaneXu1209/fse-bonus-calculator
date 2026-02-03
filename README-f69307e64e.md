# FSE奖金计算系统 - Streamlit Web应用

## 🚀 一键部署到云端（推荐）

### Streamlit Cloud 部署（免费，5分钟上线）

#### 步骤1：准备代码仓库
将以下3个文件提交到GitHub仓库（或GitLab）：
- `streamlit_app.py`（主程序）
- `requirements.txt`（依赖包）
- `README.md`（本说明文件）

#### 步骤2：部署到Streamlit Cloud

1. 访问 [Streamlit Cloud](https://share.streamlit.io/)
2. 点击 "New app" 或 "Sign in with GitHub"
3. 填写以下信息：
   - **Repository**: 选择你的GitHub仓库
   - **Branch**: `main` 或 `master`
   - **Main file path**: `streamlit_app.py`
4. 点击 "Deploy" 等待部署完成（约1-2分钟）

#### 步骤3：访问应用
部署成功后，Streamlit Cloud会提供一个访问链接，例如：
```
https://your-app-name.streamlit.app
```

#### 步骤4：分享给团队
直接将链接分享给团队成员，他们无需安装任何软件即可使用。

---

## 📋 其他部署平台

### PythonAnywhere
```bash
# 1. 上传代码到服务器
# 2. 创建虚拟环境
python3 -m venv venv
source venv/bin/activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 启动应用
streamlit run streamlit_app.py --server.port 8080
```

### HuggingFace Spaces
1. 创建新的Space，选择"Streamlit"模板
2. 上传 `streamlit_app.py` 和 `requirements.txt`
3. 自动部署完成，获得访问链接

### Render / Railway
1. 连接GitHub仓库
2. 选择Python模板
3. 自动检测并部署Streamlit应用

---

## 💻 本地运行

如果需要本地测试：

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动应用
streamlit run streamlit_app.py

# 3. 浏览器自动打开 http://localhost:8501
```

---

## 📊 功能说明

### 核心功能
- ✅ 上传FSE原始数据表和员工mapping表
- ✅ 自动执行完整的奖金计算流程
- ✅ 实时显示处理进度和结果统计
- ✅ 交互式数据展示（表格、图表）
- ✅ 一键下载计算结果（Excel格式）

### 计算范围
1. 员工名提取与匹配
2. 区域与职责信息匹配
3. 商机类型识别
4. 工程师奖金计算
5. 派工员奖金计算
6. 区域排名奖金计算
7. 后处理奖金计算

### 输出文件
- 工程师奖金表.xlsx
- 派工员奖金表.xlsx
- 区域排名奖金.xlsx
- 后处理奖金.xlsx
- FSE原始数据表_处理后.xlsx

---

## 📌 重要说明

### 派工员定义
以下职位被识别为派工员：
- Planner
- Senior Planner
- Planning Manager
- Planner - Cross Border
- Service Planning Center Supervisor

### 数据要求
- **FSE原始数据表.xlsx**: 必须包含 `Submitted By` 和 `Opportunity Name` 列
- **员工mapping表.xlsx**: 必须包含 `Name`、`Job Title`、`八大区`、`29小区` 列

### 注意事项
- 所有数据仅在当前会话中处理，不会被永久存储
- 计算结果会实时展示，可随时下载
- 支持同时处理多条记录（测试可达1000+条）

---

## 🔧 技术栈

- **框架**: Streamlit 1.28.0
- **数据处理**: Pandas 2.1.0
- **Excel读写**: openpyxl 3.1.2
- **部署**: Streamlit Cloud / PythonAnywhere / HuggingFace Spaces

---

## 📞 技术支持

如有问题，请检查：
1. 文件格式是否正确（必须是.xlsx格式）
2. 必需的列是否存在
3. 数据量是否过大（建议单次不超过10000条记录）

---

## ✨ 快速开始

1. **最简单的方式**: 使用Streamlit Cloud一键部署
2. **本地测试**: 运行 `streamlit run streamlit_app.py`
3. **分享给团队**: 发送部署后的访问链接

无需配置服务器，无需安装数据库，即开即用！
