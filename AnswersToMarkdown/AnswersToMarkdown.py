#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AnswersToMarkdown.py  (upgraded)

Microsoft Answers（日本語版）から Edge の「Web ページ、完全」保存で得られた
アーカイブ一式から、記事本文（タイトル記事）だけを抽出して Markdown に変換します。

追加機能:
1) すべてのリンク（src/href）の「スペース」を %20 に置換（URL 内の既存 %20 は維持）
2) 画像ファイルの参照が「拡張子なし」の場合、.png ファイルとして同ディレクトリにコピーし、参照も .png に書き換える
   （常時上書き。対象はローカル相対パスのみ。外部URLは対象外）

使い方:
    python3 AnswersToMarkdown.py "WebArchiveFolderName"

出力:
    WebArchiveFolderName/WebArchiveFolderName.md

依存:
    pip install beautifulsoup4
"""

import sys, os, re, shutil, urllib.parse
from pathlib import Path
from typing import Tuple, Dict
from bs4 import BeautifulSoup, NavigableString, Tag

# オプション: Markdown 出力時に "./" を削る（GitHub ではどちらでも可）
STRIP_LEADING_DOT_SLASH = True

# オプション: 拡張子なし画像をコピーするときに付ける拡張子
FORCE_COPY_NOEXT_AS = ".png"   # ご要望に合わせて PNG 固定

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

def encode_spaces_and_normalize(path_str: str) -> str:
    # 先頭 "./" を削除（任意）
    if STRIP_LEADING_DOT_SLASH and path_str.startswith("./"):
        path_str = path_str[2:]
    # スペースのみ %20 へ（既存の %20 はそのまま）
    return path_str.replace(" ", "%20")

def resolve_local_path(base_folder: Path, url_or_path: str) -> Path | None:
    # スキーム付き（http, https など）は対象外
    if re.match(r'^[a-zA-Z]+://', url_or_path or ""):
        return None
    if not url_or_path:
        return None
    # ファイルシステム上は %20 -> 空白 に戻す
    fs_path = urllib.parse.unquote(url_or_path)
    if fs_path.startswith("./"):
        fs_path = fs_path[2:]
    abs_path = (base_folder / fs_path).resolve()
    return abs_path

def ensure_png_copy_for_path(base_folder: Path, url_or_path: str, mapping: Dict[str, str]) -> Tuple[str, bool]:
    """
    - ローカル相対パスで、拡張子なし & 実在するファイルなら、同ディレクトリに .png を付けたコピーを作る
    - 以降、参照は .png に差し替え
    - 2重コピー防止のため mapping を使う
    """
    if not url_or_path or url_or_path.startswith("#"):
        return url_or_path, False
    if re.match(r'^[a-zA-Z]+://', url_or_path):
        return url_or_path, False

    key = url_or_path
    if key in mapping:
        return mapping[key], True

    abs_path = resolve_local_path(base_folder, url_or_path)
    if abs_path is None or not abs_path.exists() or abs_path.is_dir():
        return url_or_path, False

    if abs_path.suffix:
        # すでに拡張子がある場合はそのまま
        return url_or_path, False

    # .png を付けたコピーを作成
    dest_abs = abs_path.with_suffix(FORCE_COPY_NOEXT_AS)
    dest_abs.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(abs_path, dest_abs)

    # Markdown 用の相対パス（/ 区切り）
    rel = os.path.relpath(dest_abs, base_folder).replace(os.sep, "/")
    mapping[key] = rel
    return rel, True

def preprocess_links_and_images(content_root: Tag, base_folder: Path):
    """
    - <img> の src を処理（スペース -> %20、拡張子なし -> .png コピーを作って差し替え）
    - <a> の href も同様（ローカル相対パスのみ）
    """
    mapping: Dict[str, str] = {}

    # 画像（img）
    for img in content_root.find_all("img"):
        src = img.get("src")
        if not src:
            continue
        new_rel, changed = ensure_png_copy_for_path(base_folder, src, mapping)
        if changed:
            src = new_rel
        img["src"] = encode_spaces_and_normalize(src)

    # アンカー（a）
    for a in content_root.find_all("a"):
        href = a.get("href")
        if not href:
            continue
        new_rel, changed = ensure_png_copy_for_path(base_folder, href, mapping)
        if changed:
            href = new_rel
        a["href"] = encode_spaces_and_normalize(href)

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
        # 位置特定に失敗しても最低限コメントボタン自身を除去
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

def convert_html_to_markdown(html_text: str, folder_path: Path, folder_name: str) -> str:
    soup = BeautifulSoup(html_text, "html.parser")
    content_root = soup.select_one("div.thread-message-content-body-text.thread-full-message")
    if content_root is None:
        error_exit(
            "記事本文のルート要素が見つかりません。期待: "
            "<div class=\"thread-message-content-body-text thread-full-message\" ...>"
        )

    # 画像とリンクの事前処理（スペース -> %20、拡張子なし画像 -> .png コピー＆差し替え）
    preprocess_links_and_images(content_root, folder_path)

    # コメントボタン以降を削除
    strip_after_comment_button(content_root)

    # 空チェック
    if (not content_root.get_text(strip=True) and
        not content_root.find(["img", "a", "p", "ul", "ol", "table", "pre"])):
        error_exit("記事本文が空のようです。抽出範囲や保存形式を確認してください。")

    md_body = html_fragment_to_markdown(str(content_root))

    page_title = soup.title.string.strip() if soup.title and soup.title.string else folder_name
    final_md = f"# {page_title}\n\n{md_body.strip()}\n"
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

    md_text = convert_html_to_markdown(html_text, folder, folder.name)

    try:
        out_md.write_text(md_text, encoding="utf-8", newline="\n")
    except Exception as e:
        error_exit(f"Markdown の書き込みに失敗しました: {out_md}\n{e}")

    print(f"[OK] 変換完了: {out_md}")

if __name__ == "__main__":
    main()
