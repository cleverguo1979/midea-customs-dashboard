# 美的业务关务看板

华东口岸报关数据可视化看板，用于口岸单位每日上传 Excel 更新数据，领导查看业务趋势和异常跟进。

## 🌐 访问地址

| 环境 | 地址 |
|------|------|
| 公网看板 | https://cleverguo1979.github.io/midea-customs-dashboard/ |
| 本地局域网 | http://192.168.22.212:8888 |

## 🚀 部署永久云服务器（必须，只需一次）

公网看板需要后端 API 存储共享数据。以下二选一：

### 方案 A：Render 一键部署（推荐）

点击下面的按钮部署到 Render（免费套餐，24/7 运行）：

👉 **[Deploy to Render](https://render.com/deploy?repo=https://github.com/cleverguo1979/midea-customs-dashboard)**

1. 用 GitHub 账号登录 Render
2. 点击「Create Web Service」
3. 部署完成后会得到一个地址如 `https://midea-dashboard.onrender.com`
4. 把 `index.html` 中的 `API_SERVER` 常量改为这个地址
5. 提交并推送到 GitHub

### 方案 B：本机隧道（临时）

```bash
cd 美的看板
bash tunnel.sh
# 公网地址会存在 tunnel_url.txt 中
```

> 关机后隧道断开，需重新启动。

## 本地开发

```bash
pip3 install -r requirements.txt
python3 server.py
# 访问 http://localhost:8888
```

## 技术栈

- 前端：SheetJS + Chart.js + chartjs-plugin-datalabels
- 后端：Flask + openpyxl（部署在 Render）
- 数据：客户端 localStorage + 服务端 JSON 文件
