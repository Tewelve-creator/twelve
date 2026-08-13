# 面试常用单词站

包含原网页 HTML/CSS/React/Vue/JS 词表，以及新增的 **Java / Python / FastAPI / LangChain** 词条。

## 本地打开

- 直接双击 `index.html`
- 或访问本地服务：http://127.0.0.1:8766/

## 发布成公网网址（任选其一）

### Vercel（推荐）

```bash
cd Desktop/interview-vocab-site
npx vercel login
npx vercel --prod
```

发布成功后会得到类似 `https://xxx.vercel.app` 的地址。

### GitHub Pages

1. 在 GitHub 新建仓库 `interview-vocab-site`
2. 推送本目录并开启 Pages（Settings → Pages → Deploy from branch `main`）
3. 访问：`https://你的用户名.github.io/interview-vocab-site/`
