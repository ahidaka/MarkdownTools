#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AnswersToMarkdown.py  (independent-article friendly)

目的:
- Microsoft Answers の Edge 保存アーカイブから、本文を抽出し Markdown を生成
- GitHub などで「独立記事」として扱いやすいように最適化

追加機能:
1) タイトル末尾の " - Microsoft コミュニティ" を削除
2) 画像（ローカル）を Markdown 出力先フォルダへコピー
   - 拡張子なし -> .png を付けてコピー
   - .png      -> そのまま .png でコピー
   - その他の拡張子は現状対象外（必要なら拡張可能）
3) Markdown 内の画像/リンクは、コピー後の **ファイル名だけ** を参照
4) スペースを %20 に置換（URL 内の既存 %20 は維持）

使い方:
    python3 AnswersToMarkdown.py "WebArchiveFolderName"

出力:
    WebArchiveFolderName/WebArchiveFolderName.md
    + 同ディレクトリに *.png ファイル（コピー）が生成される場合あり

依存:
    pip install beautifulsoup4
"""

import sys, os, re, shutil, urllib.parse
from pathlib import Path
from typing import Tuple, Dict
from bs4 import BeautifulSoup, NavigableString, Tag

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

def local_fs_path(base_folder: Path, url_or_path: str) -> Path | None:
    if not url_or_path or is_external(url_or_path) or url_or_path.startswith("#"):
        return None
    fs_path = urllib.parse.unquote(url_or_path)
    if fs_path.startswith("./"):
        fs_path = fs_path[2:]
    return (base_folder / fs_path).resolve()

def copy_image_to_outdir(base_folder: Path, out_folder: Path, url_or_path: str, mapping: Dict[str, str]) -> Tuple[str,bool]:
    """
    ローカル相対パスの画像を Markdown 出力先にコピーし、参照をファイル名のみへ差し替える。
    - 拡張子なし -> .png でコピー
    - .png       -> そのままコピー
    - その他拡張子 -> 現状はコピーせず、そのまま（必要なら拡張可）
    """
    if not url_or_path or is_external(url_or_path) or url_or_path.startswith("#"):
        return url_or_path, False

    key = url_or_path
    if key in mapping:
        return mapping[key], True

    abs_src = local_fs_path(base_folder, url_or_path)
    if abs_src is None or not abs_src.exists() or abs_src.is_dir():
        return url_or_path, False

    if abs_src.suffix.lower() == ".png":
        dest_name = abs_src.name
    elif abs_src.suffix == "":
        dest_name = abs_src.name + ".png"
    else:
        # 非PNGは対象外（必要なら変換/コピー処理を追加）
        return url_or_path, False

    dest_abs = out_folder / dest_name
    dest_abs.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(abs_src, dest_abs)

    mapping[key] = dest_name
    return dest_name, True

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

def convert_html_to_markdown(html_text: str, base_folder: Path, out_folder: Path, folder_name: str) -> str:
    soup = BeautifulSoup(html_text, "html.parser")
    content_root = soup.select_one("div.thread-message-content-body-text.thread-full-message")
    if content_root is None:
        error_exit(
            "記事本文のルート要素が見つかりません。期待: "
            "<div class=\"thread-message-content-body-text thread-full-message\" ...>"
        )

    # 画像とリンクの事前処理（PNG系を出力先へコピー & リンクをファイル名へ）
    mapping: Dict[str, str] = {}
    for img in content_root.find_all("img"):
        src = img.get("src")
        if not src:
            continue
        new_name, changed = copy_image_to_outdir(base_folder, out_folder, src, mapping)
        if changed:
            src = new_name
        img["src"] = encode_spaces(src)

    for a in content_root.find_all("a"):
        href = a.get("href")
        if not href:
            continue
        new_name, changed = copy_image_to_outdir(base_folder, out_folder, href, mapping)
        if changed:
            href = new_name
        a["href"] = encode_spaces(href)

    # コメントボタン以降を削除
    strip_after_comment_button(content_root)

    # タイトル末尾の " - Microsoft コミュニティ" を削除
    raw_title = soup.title.string.strip() if soup.title and soup.title.string else folder_name
    stripped_title = re.sub(r"\s*-\s*Microsoft\s*コミュニティ\s*$", "", raw_title)

    # 空チェック
    if (not content_root.get_text(strip=True) and
        not content_root.find(["img", "a", "p", "ul", "ol", "table", "pre"])):
        error_exit("記事本文が空のようです。抽出範囲や保存形式を確認してください。")

    md_body = html_fragment_to_markdown(str(content_root))
    final_md = f"# {stripped_title}\n\n{md_body.strip()}\n"
    return final_md

def main():
    if len(sys.argv) != 2:
        print("使い方: python3 AnswersToMarkdown.py WebArchiveFolderName")
        sys.exit(2)

    folder = Path(sys.argv[1]).resolve()
    if not folder.exists() or not folder.is_dir():
        error_exit(f"指定フォルダが存在しません: {folder}")

    html_path = find_single_html_file(folder)
    out_md = folder / f"{folder.name}.md"

    try:
        html_text = html_path.read_text(encoding="utf-8")
    except UnicodeDecodeError as e:
        error_exit(f"HTML の UTF-8 読み込みに失敗しました: {html_path}\n{e}")

    md_text = convert_html_to_markdown(html_text, folder, folder, folder.name)

    try:
        out_md.write_text(md_text, encoding="utf-8", newline="\n")
    except Exception as e:
        error_exit(f"Markdown の書き込みに失敗しました: {out_md}\n{e}")

    print(f"[OK] 変換完了: {out_md}")

if __name__ == "__main__":
    main()
