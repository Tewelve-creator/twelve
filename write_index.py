# -*- coding: utf-8 -*-
from pathlib import Path

SITE = Path(__file__).resolve().parent / "site"
style = (SITE / "_style.css").read_text(encoding="utf-8")
script = (SITE / "_script.js").read_text(encoding="utf-8")

html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>面试常用单词卡 · 在线速查</title>
<meta name="description" content="Java / Python / FastAPI / LangChain / RAG / Transformer / OpenCV / AI 编程工具面试常用单词在线速查表">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,600;1,9..144,500&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>{style}
</style>
</head>
<body>

<header class="hero">
  <div class="kicker">Interview Vocabulary Index</div>
  <h1>面试常用单词，<em>随手一查</em>。</h1>
  <p class="lede">精选技术面试高频词：Java / Python / FastAPI / LangChain / RAG / Transformer 等八股与 API，含音标、谐音速记与一句话回答。已去掉生僻少用词，点击喇叭可听发音。</p>
  <div class="stats" id="stats">
    <div class="stat"><b id="statTotal">—</b><span>词条总数</span></div>
    <div class="stat"><b id="statTech">—</b><span>技术栈</span></div>
    <div class="stat"><b id="statMust">—</b><span>必问词条</span></div>
  </div>
</header>

<div class="wrap">
  <div class="panel">
    <div class="search-row">
      <input id="searchInput" type="text" placeholder="搜索单词、解释、中文谐音或一句话回答……" autocomplete="off" />
    </div>
    <div class="chip-groups">
      <div class="chip-group">
        <div class="label">重要程度</div>
        <div class="chips" id="levelChips"></div>
      </div>
      <div class="chip-group">
        <div class="label">技术栈</div>
        <div class="chips" id="techChips"></div>
      </div>
    </div>
  </div>

  <div class="result-bar">
    <span>共 <strong id="resultCount">0</strong> 条匹配</span>
    <button type="button" class="reset-link" id="resetBtn">清除筛选</button>
  </div>

  <div class="list" id="list"></div>
  <div class="empty" id="emptyState" style="display:none">
    <div class="glyph">∅</div>
    <div>没有找到匹配的词条，换个关键词试试。</div>
  </div>
</div>

<footer>
  <span>数据来源：面试常用单词整理表 · 共 __COUNT__ 条</span>
  <a class="dl" href="./interview-vocab.xlsx" download>下载完整 Excel 表格</a>
</footer>

<script>
{script}
</script>
</body>
</html>
"""

(SITE / "index.html").write_text(html, encoding="utf-8")
print("wrote", SITE / "index.html", len(html))
