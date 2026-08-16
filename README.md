# 面试常用单词站

搜索、筛选、音标谐音、一句话回答、朗读发音、Excel 下载。

在线地址：https://tewelve-creator.github.io/twelve/

## 用户添加单词（多人可见）

1. 打开网站，点 **管理 / 同步 Token**
2. 粘贴 GitHub Fine-grained Token（仅本仓库 **Issues: Read and write**）
3. 首页输入难词 → **AI 查找并录入**
4. AI 自动补全表头字段（分类、用法/解释、谐音、音标、发音、技术栈、重要程度、对比词、一句话回答）
5. 写入 GitHub Issue 后，Actions 自动合并进 Excel，并更新 `docs/vocab-data.json`
6. 所有人刷新页面即可看到；页面也会每 20 秒自动同步社区词条

Token 设置页快捷入口：https://tewelve-creator.github.io/twelve/?manage=1

## 本机合并到 Excel

```bash
python merge_words_to_excel.py
```

会更新：

- `e:\Users\i\Desktop\单词\面试常用单词_Java_Python_FastAPI_LangChain.xlsx`（若未被占用）
- 仓库内 Excel
- `docs/vocab-data.json`

## 本地预览

```bash
python -m http.server 8877 --directory docs
```

打开：http://127.0.0.1:8877/
