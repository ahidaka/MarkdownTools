#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AnswersToZenn.py

Edge で保存した Microsoft Answers の Web アーカイブ（HTML + *_files）から、
Zenn 形式の Markdown を生成します。

出力:
- articles/<slug>.md  （既存ならエラー停止・常に新規作成）
- images/<slug>/ ...   （本文で参照される画像をコピー）

仕様の要点:
- 本文抽出は <div class="thread-message-content-body-text thread-full-message"> から
  コメントボタン（div.message-action-container）まで。
- Markdown の先頭に Zenn Frontmatter を付与。
- title は元 HTML の <title> から末尾 " - Microsoft コミュニティ" を除去して使用。
- slug は CLI 引数で受け取り、空でないことのみ検証（既存チェックはしない）。
- 画像パスは /images/<slug>/<ファイル名> に書き換え（スペースは %20 にエンコード）。
- ローカル画像が拡張子なしなら .png を付与してコピー（中身の形式は未判定）。
- .png/.jpg/.jpeg/.gif/.webp/.bmp/.svg 等は拡張子ありのままコピー。
- 外部 URL はそのまま（スペースのみ %20 化）。
- 出力 Markdown が既に存在したらエラー終了。

依存:
  pip install beautifulsoup4
"""

import sys, os, re, shutil, urllib.parse
from pathlib import Path
from typing import Tuple, Dict, Optional, Iterable
from bs4 import BeautifulSoup, NavigableString, Tag

# ---- 固定 Frontmatter（必要時は書き換えて運用） ----
FM_EMOJI = "??"
FM_TYPE = "tech"  # "tech" or "idea"
FM_TOPICS = ["windows", "ハードウェア", "デバイスドライバー"]
FM_PUBLISHED = True

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg"}

def error_exit(msg: str, code: int = 1):
    print(f"[ERROR] {msg}", file=sys.stderr)
    sys.exit(code)

def find_single_html_file(folder: Path) -> Path:
    htmls = sorted(p for p in folder.glob("*.html") if p.is_file())
    if len(htmls) == 0:
        error_exit(f"HTML ファイルが見つかりません: {folder}")
    if len(htmls) > 1:
        names = ", ".join([h.name for h in htmls])
        error_exit(f"HTML ファイルが複数見つかりました（1つに絞ってください）: {names}")
    return htmls[0]

def is_external(url: str) -> bool:
    return bool(re.match(r'^[a-zA-Z]+://', url or ""))

def encode_spaces(s: str) -> str:
    return s.replace(" ", "%20") if s else s

def local_fs_path(base_folder: Path, url_or_path: str) -> Optional[Path]:
    if not url_or_path or is_external(url_or_path) or url_or_path.startswith("#"):
        return None
    fs_path = urllib.parse.unquote(url_or_path)
    if fs_path.startswith("./"):
        fs_path = fs_path[2:]
    return (base_folder / fs_path).resolve()

def ensure_copied_image(
    base_folder: Path,
    out_images_dir: Path,
    url_or_path: str,
    mapping: Dict[str, str]
) -> Tuple[str, bool]:
    """
    ローカル相対パスの画像を images/<slug>/ にコピーし、返り値として「/images/<slug>/<ファイル名>」を返す。
    - 拡張子なし -> .png を付与してコピー
    - 拡張子あり & IMAGE_EXTS -> その拡張子でコピー
    - 上記以外 -> 変更なし
    mapping で同一入力の重複コピーを防止。
    """
    if not url_or_path or is_external(url_or_path) or url_or_path.startswith("#"):
        return url_or_path, False

    key = url_or_path
    if key in mapping:
        return mapping[key], True

    abs_src = local_fs_path(base_folder, url_or_path)
    if abs_src is None or not abs_src.exists() or abs_src.is_dir():
        return url_or_path, False

    if abs_src.suffix.lower() in IMAGE_EXTS:
        dest_name = abs_src.name
    elif abs_src.suffix == "":
        dest_name = abs_src.name + ".png"
    else:
        # 画像拡張子ではないので対象外
        return url_or_path, False

    out_images_dir.mkdir(parents=True, exist_ok=True)
    dest_abs = out_images_dir / dest_name
    shutil.copyfile(abs_src, dest_abs)

    new_url = f"/images/{out_images_dir.name}/{dest_name}"
    mapping[key] = new_url
    return new_url, True

def strip_after_comment_button(root: Tag):
    comment = root.find("div", class_="message-action-container")
    if not comment:
        return
    parent = comment.parent
    try:
        idx = parent.contents.index(comment)
        for node in list(parent.contents)[idx:]:
            try:
                node.extract()
            except Exception:
                pass
    except Exception:
        try:
            comment.decompose()
        except Exception:
            pass

def node_to_markdown(node):
    if isinstance(node, NavigableString):
        return str(node)
    if not isinstance(node, Tag):
        return ""

    name = node.name.lower()
    if name in ["h1","h2","h3","h4","h5","h6"]:
        level = int(name[1])
        text = "".join(node_to_markdown(c) for c in node.children).strip()
        return "\n" + ("#"*level) + " " + text + "\n\n"
    if name == "p":
        text = "".join(node_to_markdown(c) for c in node.children).strip()
        return (text + "\n\n") if text else ""
    if name == "br":
        return "  \n"
    if name in ["strong","b"]:
        return "**" + "".join(node_to_markdown(c) for c in node.children) + "**"
    if name in ["em","i"]:
        return "_" + "".join(node_to_markdown(c) for c in node.children) + "_"
    if name == "a":
        href = node.get("href", "").strip()
        text = "".join(node_to_markdown(c) for c in node.children).strip() or href
        return f"[{text}]({href})" if href else text
    if name == "img":
        src = node.get("src", "").strip()
        alt = node.get("alt", "").strip()
        return f"![{alt}]({src})"
    if name == "ul":
        items = []
        for li in node.find_all("li", recursive=False):
            item = "".join(node_to_markdown(c) for c in li.children).strip()
            items.append(f"- {item}")
        return ("\n".join(items) + "\n\n") if items else ""
    if name == "ol":
        items = []
        idx = 1
        for li in node.find_all("li", recursive=False):
            item = "".join(node_to_markdown(c) for c in li.children).strip()
            items.append(f"{idx}. {item}")
            idx += 1
        return ("\n".join(items) + "\n\n") if items else ""
    if name == "code":
        text = node.get_text()
        if "\n" in text:
            return "```\n" + text + "\n```\n\n"
        else:
            return f"`{text}`"
    if name == "pre":
        return "```\n" + node.get_text() + "\n```\n\n"
    if name == "table":
        rows = node.find_all("tr")
        md = []
        for r, tr in enumerate(rows):
            cols = [c.get_text(strip=True) for c in tr.find_all(["th","td"])]
            line = "| " + " | ".join(cols) + " |"
            md.append(line)
            if r == 0:
                md.append("| " + " | ".join(["---"]*len(cols)) + " |")
        return ("\n".join(md) + "\n\n") if md else ""

    parts = [node_to_markdown(c) for c in node.children]
    return "".join(parts)

def html_fragment_to_markdown(fragment_html: str) -> str:
    frag = BeautifulSoup(fragment_html, "html.parser")
    parts = [node_to_markdown(c) for c in frag.contents]
    md = "".join(parts)
    md = re.sub(r"\n{3,}", "\n\n", md)
    md = md.strip() + "\n"
    return md

def convert_to_zenn_md(
    html_text: str,
    base_folder: Path,
    out_articles_dir: Path,
    out_images_dir: Path,
    slug: str
) -> str:
    soup = BeautifulSoup(html_text, "html.parser")
    content_root = soup.select_one("div.thread-message-content-body-text.thread-full-message")
    if content_root is None:
        error_exit(
            "記事本文のルート要素が見つかりません。期待: "
            "<div class=\"thread-message-content-body-text thread-full-message\" ...>"
        )

    # 画像・リンクの前処理（コピー & 書き換え）
    mapping: Dict[str, str] = {}
    # <img>
    for img in content_root.find_all("img"):
        src = img.get("src")
        if not src:
            continue
        new_url, changed = ensure_copied_image(base_folder, out_images_dir, src, mapping)
        if changed:
            src = new_url
        img["src"] = encode_spaces(src)
    # <a>
    for a in content_root.find_all("a"):
        href = a.get("href")
        if not href:
            continue
        new_url, changed = ensure_copied_image(base_folder, out_images_dir, href, mapping)
        if changed:
            href = new_url
        a["href"] = encode_spaces(href)

    # コメントボタン以降を削除
    strip_after_comment_button(content_root)

    # タイトル（末尾 " - Microsoft コミュニティ" を除去）
    raw_title = soup.title.string.strip() if soup.title and soup.title.string else slug
    title = re.sub(r"\s*-\s*Microsoft\s*コミュニティ\s*$", "", raw_title)

    # 本文 Markdown 生成
    md_body = html_fragment_to_markdown(str(content_root)).strip()

    # Frontmatter 生成
    topics_yaml = ", ".join([f'"{t}"' for t in FM_TOPICS])
    frontmatter = f"""---
title: "{title.replace('"', '\\"')}"
emoji: "{FM_EMOJI}"
type: "{FM_TYPE}"
topics: [{topics_yaml}]
published: {"true" if FM_PUBLISHED else "false"}
slug: "{slug}"
---

"""

    return frontmatter + md_body + "\n"

def main():
    if len(sys.argv) != 3:
        print("使い方: python3 AnswersToZenn.py WebArchiveFolderName Slug")
        sys.exit(2)

    src_folder = Path(sys.argv[1]).resolve()
    slug = sys.argv[2].strip()

    if not src_folder.exists() or not src_folder.is_dir():
        error_exit(f"指定フォルダが存在しません: {src_folder}")
    if not slug:
        error_exit("slug が空です。英数字などのユニーク ID を指定してください。")

    # 入力 HTML
    html_path = find_single_html_file(src_folder)

    # 出力先
    out_articles_dir = Path.cwd() / "articles"
    out_images_dir = Path.cwd() / "images" / slug
    out_articles_dir.mkdir(parents=True, exist_ok=True)
    out_images_dir.mkdir(parents=True, exist_ok=True)

    out_md = out_articles_dir / f"{slug}.md"
    if out_md.exists():
        error_exit(f"出力先 Markdown が既に存在します: {out_md}")

    # HTML 取得
    try:
        html_text = html_path.read_text(encoding="utf-8")
    except UnicodeDecodeError as e:
        error_exit(f"HTML の UTF-8 読み込みに失敗しました: {html_path}\n{e}")

    # 変換
    md_text = convert_to_zenn_md(html_text, src_folder, out_articles_dir, out_images_dir, slug)

    # 書き出し（UTF-8 LF）
    try:
        out_md.write_text(md_text, encoding="utf-8", newline="\n")
    except Exception as e:
        error_exit(f"Markdown の書き込みに失敗しました: {out_md}\n{e}")

    print(f"[OK] 変換完了: {out_md}")

if __name__ == "__main__":
    main()
