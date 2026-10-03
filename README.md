# 二维码打印排版网页

公开页面：<https://mzwdnmd.github.io/qr-print-sheet/>

静态网页，供客户上传带二维码的照片，在浏览器内识别并生成一页九档尺寸的 A4 Word 文档。默认基准边长 1.60 cm，间隔 0.05 cm，共 1.40–1.80 cm。参照原图，校徽直径约占二维码整体边长的 12%，周围保留直径约 23% 的白色区域；二维码使用标准字节编码和 M 级纠错。

照片和二维码内容只在当前浏览器中处理，不上传服务器。公开仓库只包含网页源码、校徽和识别程序需要的 WebAssembly 文件。

## 本地运行

```bash
npm ci
npm run dev
```

## 构建

```bash
npm run build
```

运行 `npm run build:pages` 会把可发布文件写入 `docs/`。GitHub Pages 从 `main` 分支的 `/docs` 目录发布；提交新的 `docs/` 文件后网站会更新。Vite 使用相对资源路径，可以发布在仓库子路径。打印时选择实际大小 / 100%。
