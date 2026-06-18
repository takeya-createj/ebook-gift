#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""final_book.docx -> スマホ最適化Web版（挿絵・マンガ込み）一括生成
   + 選択トップページ生成
"""
import io, re, html, hashlib, shutil, sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
from pathlib import Path
from PIL import Image
from docx import Document
from docx.oxml.ns import qn

SRC_ROOT = Path(r"G:\マイドライブ\AIコミック\最新版配布用コミクルNeo-GPT\output")
OUT_ROOT = Path(r"C:\Users\takey\雑談\電子書籍プレゼント")
BOOKS_OUT = OUT_ROOT / "books"

# 制作対象スラッグ（hongyou-fukugyo は docx 未完成のため除外）
SLUGS = [
    "suimin", "blood-sugar-sleepiness", "chatgpt-fukugyou-5man",
    "claude-gemini-fukugyou", "claude-obsidian-3x", "hongyou-wa-fukugyou",
    "nisa-ideco-guide", "obsidian-second-brain", "option-kamisama",
    "teinen-hataraki", "travel-life-impact",
]

MAX_W = 1080  # 画像最大幅(px)

# docx冒頭にタイトル見出しが無い等で自動検出できない本の補正
TITLE_OVERRIDE = {
    "claude-obsidian-3x": ("Claude Code × Obsidian 仕事3倍速メソッド",
                            "技術知識ゼロから始めるAI時代の新・仕事術"),
}

def find_cover(src_dir: Path):
    cands = ["表紙1.jpg", "表紙1.png", "表紙2.jpg", "表紙2.png"]
    for c in cands:
        p = src_dir / c
        if p.exists():
            return p
    # 変則名（例: ClaudeｰGemini副業・表紙1.jpg）
    for p in sorted(src_dir.glob("*表紙*.jpg")) + sorted(src_dir.glob("*表紙*.png")):
        return p
    return None

def save_image(img_bytes_or_path, out_dir: Path, stem_hint="img"):
    """画像を圧縮してout_dirに保存。ファイル名を返す。"""
    if isinstance(img_bytes_or_path, (bytes, bytearray)):
        im = Image.open(io.BytesIO(img_bytes_or_path))
        h = hashlib.md5(img_bytes_or_path).hexdigest()[:10]
    else:
        data = Path(img_bytes_or_path).read_bytes()
        im = Image.open(io.BytesIO(data))
        h = hashlib.md5(data).hexdigest()[:10]
    fname = f"{stem_hint}_{h}.jpg"
    out_path = out_dir / fname
    if out_path.exists():
        return fname
    im = im.convert("RGB")
    if im.width > MAX_W:
        ratio = MAX_W / im.width
        im = im.resize((MAX_W, int(im.height * ratio)), Image.LANCZOS)
    im.save(out_path, format="JPEG", quality=82, optimize=True)
    return fname

def run_html(run):
    txt = html.escape(run.text)
    if not txt:
        return ""
    is_bold = bool(run.bold)
    is_red = False
    try:
        c = run.font.color
        if c and c.rgb is not None and str(c.rgb) not in ("000000",):
            # 赤系強調（CC2200 など R>>G,B）
            rgb = str(c.rgb)
            r, g, b = int(rgb[0:2],16), int(rgb[2:4],16), int(rgb[4:6],16)
            if r > 120 and r > g + 40 and r > b + 40:
                is_red = True
    except Exception:
        pass
    if is_red:
        return f'<span class="hl">{txt}</span>'
    if is_bold:
        return f'<strong>{txt}</strong>'
    return txt

def para_images(para, img_dir):
    """段落内のインライン画像を保存してファイル名リストを返す"""
    names = []
    blips = para._element.findall('.//' + qn('a:blip'))
    doc_part = para.part
    for blip in blips:
        rId = blip.get(qn('r:embed')) or blip.get(qn('r:link'))
        if not rId:
            continue
        try:
            part = doc_part.related_parts[rId]
            names.append(save_image(part.blob, img_dir, "fig"))
        except Exception:
            pass
    return names

def build_book(slug):
    src_dir = SRC_ROOT / slug
    docx_path = src_dir / "final_book.docx"
    if not docx_path.exists():
        print(f"[SKIP] {slug}: docxなし")
        return None
    out_dir = BOOKS_OUT / slug
    img_dir = out_dir / "images"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    img_dir.mkdir(parents=True, exist_ok=True)

    # 表紙
    cover_src = find_cover(src_dir)
    cover_name = save_image(cover_src, img_dir, "cover") if cover_src else None

    doc = Document(str(docx_path))
    title, subtitle = None, None
    parts = []  # html片

    for para in doc.paragraphs:
        style = (para.style.name or "").lower()
        imgs = para_images(para, img_dir)
        if imgs:
            for nm in imgs:
                parts.append(('img', nm))
            # 画像段落にキャプションテキストが続く場合あり
            cap = para.text.strip()
            if cap:
                parts.append(('cap', html.escape(cap)))
            continue
        text = para.text.strip()
        if not text:
            continue
        # 見出し判定
        if "heading 1" in style:
            if title is None:
                title = text
                continue  # タイトルはヒーローへ
            parts.append(('h1', html.escape(text)))
        elif "heading 2" in style:
            if subtitle is None and title is not None and not any(p[0] in ('h1','p','img') for p in parts):
                subtitle = text
                continue  # サブタイトルはヒーローへ
            parts.append(('h2', html.escape(text)))
        elif "heading 3" in style:
            parts.append(('h3', html.escape(text)))
        else:
            inner = "".join(run_html(r) for r in para.runs) or html.escape(text)
            parts.append(('p', inner))

    if slug in TITLE_OVERRIDE:
        ov_t, ov_s = TITLE_OVERRIDE[slug]
        title, subtitle = ov_t, (subtitle or ov_s)
    title = title or slug
    # 本文HTML組み立て
    body = []
    toc = []
    chapter_idx = 0
    for kind, val in parts:
        if kind == 'h1':
            chapter_idx += 1
            cid = f"ch{chapter_idx:02d}"
            toc.append((cid, chapter_idx, val))
            body.append(f'<h2 class="chapter" id="{cid}"><span class="cnum">{chapter_idx:02d}</span>{val}</h2>')
        elif kind == 'h2':
            body.append(f'<h3>{val}</h3>')
        elif kind == 'h3':
            body.append(f'<h4>{val}</h4>')
        elif kind == 'p':
            body.append(f'<p>{val}</p>')
        elif kind == 'img':
            body.append(f'<figure><img loading="lazy" src="images/{val}" alt=""></figure>')
        elif kind == 'cap':
            body.append(f'<p class="caption">{val}</p>')
    body_html = "\n".join(body)

    # 目次
    toc_html = ""
    if toc:
        rows = "\n".join(
            f'<a class="toc-row" href="#{cid}"><span class="toc-num">{num:02d}</span>'
            f'<span class="toc-ttl">{ttl}</span></a>'
            for cid, num, ttl in toc
        )
        toc_html = f'<nav class="toc"><p class="toc-h">目次</p>{rows}</nav>'

    cover_html = f'<img class="cover" src="images/{cover_name}" alt="表紙">' if cover_name else ""
    sub_html = f'<p class="subtitle">{html.escape(subtitle)}</p>' if subtitle else ""

    page = BOOK_TEMPLATE.format(
        title=html.escape(title),
        subtitle_meta=html.escape(subtitle or ""),
        cover=cover_html,
        title_h=html.escape(title),
        subtitle=sub_html,
        toc=toc_html,
        body=body_html,
    )
    (out_dir / "index.html").write_text(page, encoding="utf-8")
    print(f"[OK] {slug}: {title} / 画像{len(list(img_dir.glob('*.jpg')))}枚")
    return {"slug": slug, "title": title, "subtitle": subtitle or "", "cover": cover_name}

BOOK_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{title}</title>
<meta name="description" content="{subtitle_meta}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{{--ink:#23201c;--sub:#6f675c;--line:#e7e0d4;--bg:#f7f3ec;--card:#fffdf8;--accent:#c0392b;--brand:#3a5a8c;}}
*{{box-sizing:border-box;}}
html{{scroll-behavior:smooth;}}
body{{margin:0;background:var(--bg);color:var(--ink);font-family:"Noto Serif JP",serif;line-height:1.95;font-size:17px;-webkit-text-size-adjust:100%;}}
.bar{{position:fixed;top:0;left:0;height:3px;width:0;background:linear-gradient(90deg,var(--brand),var(--accent));z-index:50;transition:width .1s;}}
.wrap{{max-width:680px;margin:0 auto;padding:0 22px 80px;}}
header.hero{{text-align:center;padding:46px 0 30px;}}
.cover{{width:min(70%,300px);height:auto;border-radius:10px;box-shadow:0 18px 40px rgba(60,40,10,.22);margin-bottom:26px;}}
.hero h1{{font-size:26px;font-weight:700;line-height:1.5;margin:.2em 0;letter-spacing:.01em;}}
.subtitle{{color:var(--sub);font-size:15px;font-family:"Noto Sans JP",sans-serif;margin:.4em auto 0;max-width:90%;}}
.divider{{width:54px;height:3px;background:var(--accent);border:0;border-radius:3px;margin:26px auto 8px;}}
.toc{{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:10px 18px 16px;margin:8px 0 10px;box-shadow:0 6px 18px rgba(60,40,10,.07);}}
.toc-h{{font-family:"Noto Sans JP",sans-serif;font-weight:700;font-size:15px;letter-spacing:.18em;color:var(--brand);text-align:center;margin:12px 0 14px;}}
.toc-row{{display:flex;align-items:baseline;gap:13px;padding:11px 4px;text-decoration:none;color:var(--ink);border-bottom:1px dashed var(--line);font-family:"Noto Sans JP",sans-serif;line-height:1.5;}}
.toc-row:last-child{{border-bottom:0;}}
.toc-row:active{{background:#f1ebe0;}}
.toc-num{{flex:0 0 auto;font-size:13px;font-weight:700;color:var(--accent);min-width:1.6em;}}
.toc-ttl{{font-size:14.5px;}}
.chapter{{font-family:"Noto Sans JP",sans-serif;font-weight:700;font-size:21px;line-height:1.45;margin:62px 0 22px;padding-top:26px;border-top:1px solid var(--line);color:var(--brand);scroll-margin-top:14px;}}
.chapter .cnum{{display:block;font-size:13px;letter-spacing:.22em;color:var(--accent);margin-bottom:6px;}}
h3{{font-family:"Noto Sans JP",sans-serif;font-weight:700;font-size:18px;margin:40px 0 14px;padding-left:13px;border-left:5px solid var(--brand);line-height:1.5;}}
h4{{font-family:"Noto Sans JP",sans-serif;font-weight:500;font-size:16px;color:var(--sub);margin:28px 0 10px;}}
p{{margin:0 0 1.25em;}}
.hl{{color:var(--accent);font-weight:700;}}
strong{{font-weight:700;}}
figure{{margin:30px 0;text-align:center;}}
figure img{{width:100%;height:auto;border-radius:10px;box-shadow:0 8px 22px rgba(60,40,10,.13);}}
.caption{{font-family:"Noto Sans JP",sans-serif;font-size:13px;color:var(--sub);text-align:center;margin:-18px 0 24px;}}
.backtop{{display:block;text-align:center;margin:60px auto 0;font-family:"Noto Sans JP",sans-serif;font-size:14px;color:var(--brand);text-decoration:none;border:1px solid var(--line);border-radius:30px;padding:12px 0;max-width:280px;background:var(--card);}}
.foot{{text-align:center;color:var(--sub);font-size:12px;font-family:"Noto Sans JP",sans-serif;margin-top:40px;}}
@media(max-width:480px){{body{{font-size:16px;}}.hero h1{{font-size:22px;}}.chapter{{font-size:19px;}}}}
</style>
</head>
<body>
<div class="bar" id="bar"></div>
<div class="wrap">
<header class="hero">
{cover}
<h1>{title_h}</h1>
{subtitle}
<hr class="divider">
</header>
{toc}
<main>
{body}
<a class="backtop" href="../../index.html">&#9664; 他の書籍を見る</a>
<p class="foot">© 竹谷知悦 / 特別プレゼント版</p>
</main>
</div>
<script>
var bar=document.getElementById('bar');
addEventListener('scroll',function(){{
 var h=document.documentElement.scrollHeight-innerHeight;
 bar.style.width=(h>0?(scrollY/h*100):0)+'%';
}});
</script>
</body>
</html>"""

def build_index(books):
    cards = []
    for b in books:
        cov = f'books/{b["slug"]}/images/{b["cover"]}' if b["cover"] else ""
        img = f'<img loading="lazy" src="{cov}" alt="">' if cov else '<div class="noimg"></div>'
        sub = f'<p class="csub">{html.escape(b["subtitle"])}</p>' if b["subtitle"] else ""
        title_disp = html.escape(b["title"])
        title_attr = html.escape(b["title"].replace("\n", " "), quote=True)
        cards.append(f'''<a class="card" data-slug="{b["slug"]}" data-title="{title_attr}" href="books/{b["slug"]}/index.html">
<div class="thumb">{img}<span class="lockbadge">🔒</span></div>
<div class="cmeta"><h3>{title_disp}</h3>{sub}<span class="readtag">▶ この本を読む</span></div>
</a>''')
    page = INDEX_TEMPLATE.replace("__CARDS__", "\n".join(cards)).replace("__N__", str(len(books)))
    (OUT_ROOT / "index.html").write_text(page, encoding="utf-8")
    print(f"[OK] index.html ({len(books)}冊)")

INDEX_TEMPLATE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>電子書籍プレゼント｜お好きな1冊をどうぞ</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--ink:#23201c;--sub:#6f675c;--line:#e7e0d4;--bg:#f7f3ec;--card:#fffdf8;--accent:#c0392b;--brand:#3a5a8c;}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--ink);font-family:"Noto Sans JP",sans-serif;}
.head{text-align:center;padding:54px 22px 8px;}
.badge{display:inline-block;font-size:12px;letter-spacing:.2em;color:#fff;background:var(--accent);padding:6px 16px;border-radius:30px;}
.head h1{font-family:"Noto Serif JP",serif;font-size:25px;font-weight:700;margin:18px 0 8px;line-height:1.5;}
.head p{color:var(--sub);font-size:14px;margin:0;}
.wrap{max-width:760px;margin:0 auto;padding:24px 18px 70px;}
#banner{display:none;background:linear-gradient(135deg,var(--accent),#a02a1c);color:#fff;text-align:center;padding:14px 16px;border-radius:13px;margin:0 0 20px;font-size:14px;line-height:1.6;box-shadow:0 6px 18px rgba(192,57,43,.25);}
#banner b{font-weight:700;font-size:15px;}
#banner .sub{font-size:12px;opacity:.92;display:block;margin-top:3px;}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;}
.card{position:relative;background:var(--card);border:1px solid var(--line);border-radius:14px;overflow:hidden;text-decoration:none;color:inherit;display:flex;flex-direction:column;box-shadow:0 4px 14px rgba(60,40,10,.07);transition:transform .15s,box-shadow .15s,opacity .25s,filter .25s;}
.card:active{transform:scale(.98);}
.thumb{position:relative;}
.card img,.noimg{width:100%;aspect-ratio:3/4;object-fit:cover;display:block;background:#ece5d8;}
.cmeta{padding:13px 13px 16px;}
.cmeta h3{font-size:14.5px;font-weight:700;margin:0 0 5px;line-height:1.45;}
.csub{font-size:11.5px;color:var(--sub);margin:0 0 8px;line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.readtag{font-family:"Noto Sans JP",sans-serif;font-size:12px;font-weight:700;color:var(--brand);}
.lockbadge{display:none;position:absolute;inset:0;align-items:center;justify-content:center;flex-direction:column;gap:4px;font-size:30px;background:rgba(35,30,20,.42);color:#fff;}
.lockbadge::after{content:"選択済み";font-size:11px;font-weight:700;letter-spacing:.05em;}
.card.locked{opacity:.5;filter:grayscale(.7);}
.card.locked .lockbadge{display:flex;}
.card.locked .readtag{display:none;}
.card.chosen{outline:3px solid var(--accent);outline-offset:-1px;box-shadow:0 10px 26px rgba(192,57,43,.3);}
.card.chosen .readtag{color:var(--accent);}
.card.chosen .readtag::before{content:"";}
.foot{text-align:center;color:var(--sub);font-size:12px;margin-top:46px;}
@media(min-width:620px){.grid{grid-template-columns:repeat(3,1fr);}}
</style>
</head>
<body>
<div class="head">
<span class="badge">SPECIAL GIFT</span>
<h1>お好きな1冊をプレゼント</h1>
<p id="lead">ご登録ありがとうございます。気になる1冊をタップしてお選びください（全__N__冊）</p>
</div>
<div class="wrap">
<div id="banner"></div>
<div class="grid">
__CARDS__
</div>
<p class="foot">© 竹谷知悦 / LINE登録特典</p>
</div>
<script>
(function(){
  var KEY='ebookgift_choice_v1';
  var p=new URLSearchParams(location.search);
  if(p.has('reset')){ try{localStorage.removeItem(KEY);}catch(e){} history.replaceState({},'',location.pathname); }
  var chosen=null; try{chosen=localStorage.getItem(KEY);}catch(e){}
  var cards=[].slice.call(document.querySelectorAll('.card'));
  var banner=document.getElementById('banner');
  var lead=document.getElementById('lead');
  function titleOf(slug){var t='';cards.forEach(function(c){if(c.getAttribute('data-slug')===slug)t=c.getAttribute('data-title');});return t;}
  function applyLock(){
    cards.forEach(function(c){
      if(c.getAttribute('data-slug')===chosen){c.classList.add('chosen');}
      else{c.classList.add('locked');}
    });
    if(banner){
      banner.innerHTML='🎁 あなたが選んだ本<br><b>'+titleOf(chosen)+'</b><span class="sub">タップして続きを読む</span>';
      banner.style.display='block';
    }
    if(lead){lead.textContent='プレゼントの1冊を選択済みです。下のあなたの本をお読みください。';}
  }
  cards.forEach(function(c){
    c.addEventListener('click',function(e){
      var slug=c.getAttribute('data-slug'), title=c.getAttribute('data-title');
      if(chosen){
        if(slug!==chosen){ e.preventDefault(); alert('すでに「'+titleOf(chosen)+'」をお選びいただいています。\\nプレゼントは1冊だけです。'); }
        return;
      }
      e.preventDefault();
      if(confirm('「'+title+'」を受け取りますか？\\n\\n※プレゼントは1冊だけ。あとから選び直しはできません。')){
        try{localStorage.setItem(KEY,slug);}catch(e){}
        location.href=c.getAttribute('href');
      }
    });
  });
  if(chosen) applyLock();
})();
</script>
</body>
</html>"""

if __name__ == "__main__":
    books = []
    for slug in SLUGS:
        r = build_book(slug)
        if r:
            books.append(r)
    build_index(books)
    print(f"\n=== 完了: {len(books)}冊 + 選択トップ ===")
    print(f"出力先: {OUT_ROOT}")
