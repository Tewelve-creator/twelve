# 面试常用单词站

样式对齐 [Netlify 示例站](https://roaring-alpaca-e98168.netlify.app/)：搜索、重要程度/技术栈筛选、音标谐音、一句话回答、朗读发音、Excel 下载。

## 本地预览

```bash
python build_site.py
python -m http.server 8877 --directory docs
```

打开：http://127.0.0.1:8877/

## 发布到 GitHub Pages

1. 把本仓库推到 GitHub（公开仓库）
2. 仓库 **Settings → Pages → Build and deployment**：
   - Source 选 **GitHub Actions**  
   或 Source 选 **Deploy from a branch**，Branch=`master`，Folder=`/docs`
3. 推送后等待 Actions 成功，访问：

`https://<你的用户名>.github.io/<仓库名>/`

静态文件在 [`docs/`](docs/)：`index.html`、`vocab-data.json`、`interview-vocab.xlsx`。

## 重新从 Excel 生成

```bash
python build_site.py
```
