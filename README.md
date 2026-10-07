# UL 实验室设备管理

面向实验室设备借用与归还的响应式网页应用，支持电脑和手机访问。

## 在线网站

- Cloudflare Pages：<https://ul-lab-equipment-app.pages.dev/>
- Vercel 备用地址：<https://ul-lab-equipment-app.vercel.app/>

## 已包含内容

- React / Vite 网页源码与页面样式
- UL Solutions 标识、二维码和页面图片资源
- Supabase 云端同步逻辑（环境变量不入库）
- 借用、归还、记录筛选、状态提示及 Excel 导出功能
- 可下载的 Excel 自动借还管理工作簿

## 本地运行

```bash
npm install
npm run dev
```

## 部署

Cloudflare Pages 连接本仓库的 `main` 分支后，会在每次推送后自动构建部署。

构建命令：`npm run build`  
输出目录：`dist`

> `.env` 中的 Supabase 配置及云端业务数据不会提交到 GitHub。
