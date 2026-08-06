# Vigor 产品编码器

基于 API 标准（石油钻采行业）的产品编码生成与管理系统。

- 编码规则：标准（A~Z）/ 子类（A~Z）/ 产品（AA~ZZ）/ 关键参数（最多 7 个）
- 支持：数据管理、编码生成、CRM 批量导出、Excel 批量导入
- 两种部署形态：本地版（局域网共享）+ 云端版（腾讯云 CloudBase）

## 目录结构

```
├── local/                 # 本地版（Python 标准库，零依赖可跑）
│   ├── server.py          #   HTTP 服务（端口 8765，数据存本机 data.json）
│   ├── 启动编码器.bat       #   Windows 一键启动
│   ├── public/            #   前端页面 + 导入模板
│   └── template/          #   CRM 导入模板
├── cloud/                 # 云端版（CloudBase 云托管容器型）
│   ├── server.py          #   云端服务（数据持久化到云存储 COS）
│   ├── Dockerfile
│   └── requirements.txt   #   openpyxl + cos-python-sdk-v5
├── web/                   # 入口跳转页（CloudBase 静态托管）
└── standalone/            # 单文件版（整个应用一个 HTML，数据存浏览器）
```

## 本地版使用

```bash
cd local
python server.py
# 浏览器访问 http://127.0.0.1:8765 （局域网内其他电脑用 http://本机IP:8765）
```

数据保存在 `local/data.json`（可备份/迁移），上传文件在 `local/uploads/`。

## 云端版部署（腾讯云 CloudBase 云托管）

1. 环境变量配置（云托管服务 EnvParams）：
   - `TCB_SECRETID` / `TCB_SECRETKEY`：腾讯云 API 密钥（用于访问云存储持久化数据）
   - 数据对象存储路径：`vigor/data.json`（bucket 需与代码中 `COS_BUCKET` 一致）
2. 容器监听端口：`PORT`（默认 3000）
3. 部署：`manageCloudRun(action="deploy", serverType="container")` 或控制台上传

## 编码规则

| 段 | 规则 | 容量 |
|----|------|------|
| 标准 | A ~ Z（单字母） | 26 |
| 子类 | A ~ Z（标准内单字母） | 26/标准 |
| 产品 | AA ~ ZZ（子类内双字母） | 676/子类 |
| 参数 | K1 ~ K7（最多 7 组，每组编码+描述） | - |

编码自动生成，无需人工指定。

## 功能

- **编码器**：逐步选择 标准 → 子类 → 产品 → 参数，自动生成完整编码
- **数据管理**：标准/子类/产品/参数组/选项的增删改，多人共享实时同步
- **CRM 导出**：按模板生成 7 列（产品编码/分类/单位/启用/中英描述/备注）导入文件
- **批量导入**：Excel 标准表格（只填中英文名，编码自动），支持最大 7 参数
- **附件管理**：产品可挂附件（图纸/证书），本地版存本机，云端版存云存储

## 备注

- `data.json`、`uploads/` 为运行时数据，不入库（已在 .gitignore）
- 产品参数必须基于真实行业数据（API 标准），禁止编造
