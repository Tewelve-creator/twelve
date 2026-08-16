# 面试常用单词站

搜索、筛选、音标谐音、一句话回答、朗读发音、Excel 下载。

在线地址：https://tewelve-creator.github.io/twelve/

## 使用（无需 Token）

1. 打开网站，输入难词 → **AI 查找并录入**
2. AI 自动补全表头字段（分类、用法/解释、谐音、音标、发音、技术栈、重要程度、对比词、一句话回答）
3. 词条立刻出现在当前页面，并邮件通知维护者（`3437903275@qq.com`）
4. 页面仍会同步已写入仓库的社区词条（`community-words.json` / Excel）

访客与维护者都不必在网页里粘贴 GitHub Token。

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
