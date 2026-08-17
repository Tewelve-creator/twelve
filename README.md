# 面试常用单词站

搜索、筛选、音标谐音、一句话回答、朗读发音、Excel 下载。

在线地址：https://tewelve-creator.github.io/twelve/

## 使用

1. 输入难词 → **AI 查找并录入**
2. 若已开启多人同步：自动写入，其他人刷新可见
3. 若未开启：会打开 GitHub 预填页，点一次 **Submit new issue** 后全员可见

## 开启「任何人添加，自动同步」（推荐做一次）

你已配置 `SUBMIT_PAT` 的话，还差一步：**让网站用 Actions 发布**，否则线上 Token 一直是空的，别人加的词同步不了。

1. 打开：https://github.com/Tewelve-creator/twelve/settings/pages  
2. **Build and deployment → Source** 选 **GitHub Actions**（不要选 “Deploy from a branch”）  
3. 打开：https://github.com/Tewelve-creator/twelve/actions/workflows/pages.yml → **Run workflow**  
4. 部署成功后，用无痕窗口打开站点试加一个词；不应再弹出 GitHub 提交页

Secret 名必须是 `SUBMIT_PAT`：https://github.com/Tewelve-creator/twelve/settings/secrets/actions

## 网站添加 → 自动写入 Excel

添加成功后流程：

1. 站点创建带 `word-submission` 标签的 Issue  
2. GitHub Actions 运行 `Sync community words`  
3. 写入仓库 Excel 工作表 **「社区添加单词」**，并同步到 **「全部汇总」**  
4. 更新 `docs/vocab-data.json` 与下载用 `interview-vocab.xlsx`

也可本机手动合并：

```bash
python merge_words_to_excel.py
```


## 本地预览

```bash
python -m http.server 8877 --directory docs
```
