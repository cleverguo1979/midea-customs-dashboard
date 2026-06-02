# 美的业务关务看板

华东口岸报关数据可视化看板，用于口岸单位每日上传 Excel 更新数据，领导查看业务趋势和异常跟进。

## 生产环境

公网地址：**https://cleverguo1979.github.io/midea-customs-dashboard/**

- 首次打开为空，需口岸单位上传 Excel 报表
- 上传后数据保存在浏览器 localStorage 中，关闭/刷新不丢失
- 每次上传**完全替换**旧数据（最新一份为唯一数据源）
- 不同设备独立存储，互不影响

## 数据格式

Excel 文件需包含以下 Sheet：
- `XXXX 日报关情况` — 日期、应申报总单量、已放行单量、审结、未放行原因
- `XXXX年异常跟踪表` — 异常记录明细
- `Sheet3` — 经营单位 → 事业部 → 通关代表 映射表

## 本地开发

```bash
cd 美的看板
python3 server.py
# 访问 http://localhost:8888
```

本地测试时可放置任意 `.xlsx` 文件在项目目录下用于调试。

## 技术栈

- SheetJS (xlsx) — Excel 解析
- Chart.js — 图表渲染
- chartjs-plugin-datalabels — 图表数据标签
- 纯前端，无后端依赖
