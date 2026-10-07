# 信息提取器 SaaS MVP

独立 Next.js MVP：把文本、图片/截图、PDF、Excel 资料源，结合文字要求、Excel 模板或截图模板，提取为文本结果和 Excel 表格。

## 功能范围

- 简单密码保护：默认开发密码 `demo123`，生产环境用 `APP_PASSWORD` 覆盖。
- 输入源：文本、图片/截图、PDF、Excel `.xlsx`。
- PDF：浏览器端提取文字并渲染页面图片；MVP 最多 20 页，超过会提示拆分。
- 模板约束：文字要求、Excel 模板表头解析、截图模板图片。
- 输出：文本结果、表格预览、复制 TSV、下载 `.xlsx`。
- AI：默认 mock 可用；配置服务端 env 后走真实接口。

## 服务端环境变量

```bash
APP_PASSWORD=your-password
CHATGPT_API_URL=https://example.com/codex/v1/responses
CHATGPT_API_KEY=replace-with-server-only-key
CHATGPT_MODEL=gpt-5.5
```

不要把真实 key 写入代码、README、日志或浏览器 bundle。

## 开发

```bash
npm install
npm run dev
```

## 验证

```bash
npm run lint
npm run typecheck
npm run build
npm audit --omit=dev
```
