# 美的项目客服看板 v2

华东口岸报关数据可视化管理平台。支持 Excel 上传建立基线 + 网页实时登记、跟进、闭环管理。

## 🆕 v2 新功能

- **实时登记**：上传一次 Excel 后，所有后续业务登记在网页上直接完成
- **逐条编辑**：每条异常记录可修改状态、进度、责任方等
- **跟进备注**：每条记录支持添加多条跟进备注，时间线展示
- **批量操作**：勾选多条记录，一键批量标记已闭环/未闭环
- **业务量登记**：手动登记每日申报数据
- **操作日志**：所有增删改操作永久记录，可追溯
- **二次确认清零**：清零需两次输入「确认清零」防止误操作
- **SQLite 存储**：从 JSON 文件升级为 SQLite 数据库，并发安全、不丢数据

## 🌐 访问地址

| 环境 | 地址 |
|------|------|
| 公网看板 | https://cleverguo1979.github.io/midea-customs-dashboard/ |
| 本地局域网 | http://192.168.22.212:8888 |

## 🚀 部署

### 方案 A：Render 一键部署（推荐）

1. 用 GitHub 账号登录 [Render](https://render.com)
2. 创建 Web Service，连接此仓库
3. Build Command: `pip install flask openpyxl gunicorn`
4. Start Command: `python server.py`
5. 部署后得到地址如 `https://midea-dashboard.onrender.com`
6. 修改 `index.html` 中 `API_SERVER` 为你的 Render 地址

### 方案 B：本机运行

```bash
cd 美的看板
pip3 install -r requirements.txt
python3 server.py
# 访问 http://localhost:8888
```

## 📋 使用指南

### 初始数据导入
1. 点击「上传数据」按钮
2. 选择 Excel 报表文件（.xlsx/.xls）
3. 系统自动解析并导入（merge 模式，不覆盖手动数据）

### 日常操作
- **登记业务量**：点击导航栏「📊 登记业务量」
- **新增异常**：点击导航栏「＋ 新增登记」
- **编辑记录**：点击表格行中的「✏️」按钮
- **添加备注**：点击「💬」按钮，在编辑面板底部添加
- **切换状态**：直接点击表格中的「已闭环/未闭环」标签
- **批量操作**：勾选行首复选框，使用批量工具栏
- **查看日志**：点击「📜 操作日志」

### 数据安全
- 所有操作实时保存到服务器数据库
- 刷新页面/关闭浏览器不会丢失数据
- 清零需要两次输入「确认清零」
- 操作日志永久记录所有变更

## 技术栈

- 前端：原生 JS + SheetJS + Chart.js + chartjs-plugin-datalabels
- 后端：Flask + SQLite + openpyxl
- 存储：SQLite（从 v1 的 JSON 文件升级）
- 部署：GitHub Pages（前端）+ Render（后端）

## 数据迁移

首次启动 v2 时，系统自动检测 `data.json` 并迁移到 SQLite，完成后备份为 `data.json.bak`。
