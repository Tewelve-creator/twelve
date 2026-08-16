# 面试常用单词站

搜索、筛选、音标谐音、一句话回答、朗读发音、Excel 下载。

在线地址：https://tewelve-creator.github.io/twelve/

## 使用（访客无需填 Token）

1. 打开网站，输入难词 → **AI 查找并录入**
2. AI 自动补全表头字段并立刻显示在页面上
3. 若维护者已配置 `SUBMIT_PAT`，会自动创建 GitHub Issue → Actions 合并进 Excel / `community-words.json`
4. 其他人打开或刷新页面（约每 20 秒自动同步）即可看到新词

## 维护者：一次配置，实现多人自动同步

仓库 Settings → Secrets and variables → Actions → New repository secret：

| Name | Value |
|------|--------|
| `SUBMIT_PAT` | Fine-grained PAT，仅本仓库 **Issues: Read and write** |

然后推送 `docs/` 或手动运行工作流 **Deploy GitHub Pages**。部署时会把 Token 写入线上站点配置（**不提交进 git**）。

未配置时：仍可 AI 录入 + 邮件通知 `3437903275@qq.com`，但不会自动建 Issue，其他人暂时看不到。

## 本机合并到 Excel

```bash
python merge_words_to_excel.py
```

## 本地预览

```bash
python -m http.server 8877 --directory docs
```

打开：http://127.0.0.1:8877/
