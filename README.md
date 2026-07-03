<div align="center">

<img src="icons/icon-192.png" width="88" alt="轻转 LiteConvert">

# 轻转 LiteConvert

**本地离线的文件格式转换工具 · Word / PPT / PDF / 图片互转**

所有转换都在你的浏览器里完成 —— 文件永不上传,没有服务器,断网也能用。

[在线使用](#在线使用) · [安装到 Mac / iPhone](#安装到设备) · [本地运行](#本地运行) · [English](#english)

</div>

---

## ✨ 功能

| 输入 | 可转换为 |
| --- | --- |
| 图片 (PNG / JPG / WebP / GIF / BMP / TIFF / HEIC*) | PNG · JPG · WebP · BMP · TIFF · GIF · PDF(支持多张合并) |
| PDF | 图片 (PNG / JPG) · Word (.docx) · PPT (.pptx) · 纯文本 |
| Word (.docx) | PDF · 纯文本 |
| PPT (.pptx) | PDF · 图片 (PNG) |

\* HEIC / TIFF 的读取依赖系统解码,推荐在 Safari(Mac / iPhone)中使用。

> **说明**:PDF → Word / PPT 属于业界难题。本工具会尽力提取文本与版面(PDF → PPT 采用整页高清图片方案,观感 100% 一致但文字不可编辑;PDF → Word 提取可编辑文本但排版会简化)。

## 🔒 隐私

- 转换引擎 100% 运行在浏览器本地(WebAssembly / Canvas / 纯 JS)
- 无任何网络请求、统计或埋点,文件从不离开你的设备
- 代码完全开源,欢迎审查

## 📲 安装到设备

**Mac(Safari)**:打开网页 → 菜单栏「文件」→「添加到程序坞…」,即可像 App 一样使用,支持离线。

**iPhone / iPad(Safari)**:打开网页 → 分享按钮 →「添加到主屏幕」。

**其他浏览器(Chrome / Edge)**:地址栏右侧的「安装」图标。

## 💻 本地运行

无需构建、无需依赖,克隆即用:

```bash
git clone https://github.com/<你的用户名>/liteconvert.git
cd liteconvert
python3 serve.py        # 打开 http://127.0.0.1:8973
```

macOS 用户也可以直接双击 `启动.command`。

## 🚀 部署自己的版本

这是一个纯静态站点,任何静态托管都可以:

- **GitHub Pages**:仓库 Settings → Pages → Branch 选 `main` → 保存,几分钟后即可通过 `https://<用户名>.github.io/liteconvert/` 访问
- 也可部署到 Cloudflare Pages、Vercel、Netlify 等

## 🧱 技术栈

- 原生 HTML / CSS / JavaScript,无构建步骤
- [pdf.js](https://mozilla.github.io/pdf.js/) — PDF 渲染与文本提取
- [pdf-lib](https://pdf-lib.js.org/) — PDF 生成
- [docx](https://docx.js.org/) — Word 文档生成
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) — PPT 生成
- [JSZip](https://stuk.github.io/jszip/) — OOXML 解析与打包
- 自研引擎:Word / PPT 版面渲染(OOXML → Canvas)、BMP / TIFF / GIF 编码器

## 🧪 测试

```bash
node test/run_tests.mjs   # 需要 Node.js ≥ 18 与 Playwright
```

17 个端到端用例覆盖全部转换方向,输出经 Pillow / pypdf / python-docx / python-pptx 独立校验。

## 📄 许可

[MIT](LICENSE) — 自由使用、修改、分发。

---

## English

**LiteConvert** is a free, open-source, fully offline file format converter for Word, PowerPoint, PDF and images. Every conversion runs locally in your browser — no uploads, no servers, works offline. Install it as a PWA on macOS ("Add to Dock" in Safari) or iOS ("Add to Home Screen"), or self-host it anywhere as a static site. See the Chinese sections above for the full feature matrix; the UI is currently in Chinese, and localization PRs are welcome.

MIT licensed.
