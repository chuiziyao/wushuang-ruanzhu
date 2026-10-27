# -*- coding: utf-8 -*-
"""生成软件著作权「程序鉴别材料」Word。

产物：
  1. 源代码_完整.docx      全部自研源代码，每页固定行数、无空行、右上角页码
  2. 源代码_前30后30.docx  仅当总页数 > 60 时生成；前 30 页 + 后 30 页

硬约束（用户与审查口径）：
  · 每页 ≥ 50 行（默认 50 行/页，可用 --lines-per-page 调整）
  · 全文无空行
  · 页眉左：软件全称 + 版本号；右：第 X 页 共 N 页
  · 页码与页数由生成器本身确定，不依赖 Word 事后排版

用法：
  python gen_source_docx.py --root D:/proj --name "示例事项软件" --version V1.0 --outdir out
"""
import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from docx_common import (  # noqa: E402
    new_document, setup_page, build_header, set_page_number_start,
    clone_sectpr_into_last_paragraph, set_run_font, set_update_fields,
)
from docx.shared import Pt  # noqa: E402
import config as cfgmod  # noqa: E402

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

EXTS = {
    '.py': 'Python', '.js': 'JavaScript', '.mjs': 'JavaScript', '.cjs': 'JavaScript',
    '.ts': 'TypeScript', '.tsx': 'TypeScript', '.jsx': 'JavaScript',
    '.vue': 'Vue', '.kt': 'Kotlin', '.java': 'Java', '.go': 'Go',
    '.c': 'C', '.h': 'C', '.cpp': 'C++', '.cc': 'C++', '.hpp': 'C++',
    '.cs': 'C#', '.swift': 'Swift', '.rs': 'Rust', '.php': 'PHP',
    '.rb': 'Ruby', '.m': 'Objective-C', '.dart': 'Dart', '.lua': 'Lua',
    '.ets': 'ArkTS', '.sql': 'SQL', '.html': 'HTML', '.htm': 'HTML',
    '.css': 'CSS', '.scss': 'SCSS', '.less': 'LESS', '.sh': 'Shell',
    '.ps1': 'PowerShell', '.xml': 'XML', '.json': 'JSON', '.yml': 'YAML',
    '.yaml': 'YAML', '.gradle': 'Gradle', '.toml': 'TOML', '.ini': 'INI',
    '.mm': 'Objective-C', '.asm': 'Assembly', '.s': 'Assembly',
}

SKIP_DIRS = {
    'node_modules', '.git', '.svn', '.hg', 'dist', 'build', 'out', 'target',
    '.gradle', '.idea', '.vscode', '__pycache__', 'venv', '.venv', 'env',
    'site-packages', 'vendor', 'bower_components', 'coverage', '.next', '.nuxt',
    '.workbuddy', '.claude', 'Library', 'Temp', 'Logs', 'obj', 'bin',
    'Pods', 'DerivedData', '.dart_tool', '.pub-cache', 'release', 'debug',
    'third_party', 'thirdparty', 'static_lib', 'jniLibs',
}

SKIP_FILE_RE = re.compile(
    r'(package-lock\.json|yarn\.lock|pnpm-lock\.yaml|poetry\.lock|Pipfile\.lock|'
    r'Cargo\.lock|go\.sum|composer\.lock|\.min\.(js|css)$|\.map$|\.pyc$|'
    r'\.(png|jpg|jpeg|gif|ico|svg|webp|bmp|ttf|otf|woff2?|eot|mp3|mp4|wav|zip|'
    r'rar|7z|pdf|docx?|xlsx?|pptx?|exe|dll|so|dylib|apk|ipa|jar|class|o|a|'
    r'bin|dat|db|sqlite|lock)$)', re.I)

SECRET_RE = re.compile(
    r'(secret|credential|passwd|password|\.env|id_rsa|\.pem$|\.pfx$|'
    r'token|apikey|api_key|llm_secret|private_key)', re.I)


def display_width(text):
    """按显示宽度计（中文/全角算 2，其余算 1），用于控制换行不溢出。"""
    w = 0
    for ch in text:
        w += 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
    return w


def normalize_line(line):
    line = line.replace('\r', '').replace('\u00a0', ' ').replace('\t', '    ')
    line = ''.join(ch for ch in line if ch == ' ' or ord(ch) >= 32)
    return line.rstrip()


def read_text(path):
    for enc in ('utf-8', 'utf-8-sig', 'gb18030'):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except (UnicodeDecodeError, OSError):
            continue
    return None


def collect_files(root, extra_skip_dirs=None, extra_skip_re=None):
    skip_dirs = set(SKIP_DIRS) | set(extra_skip_dirs or [])
    kept, skipped = [], {'secret': [], 'minified': [], 'large': [], 'noext': 0}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames
                             if d not in skip_dirs and not d.startswith('.'))
        for fn in sorted(filenames):
            ext = os.path.splitext(fn)[1].lower()
            if ext not in EXTS:
                continue
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, root).replace('\\', '/')
            if SKIP_FILE_RE.search(fn):
                continue
            if extra_skip_re and extra_skip_re.search(rel):
                continue
            if SECRET_RE.search(fn) or SECRET_RE.search(rel):
                skipped['secret'].append(rel)
                continue
            try:
                size = os.path.getsize(full)
            except OSError:
                continue
            if size > 2 * 1024 * 1024:
                skipped['large'].append(rel)
                continue
            text = read_text(full)
            if text is None:
                continue
            raw_lines = text.split('\n')
            longest = max((len(l) for l in raw_lines[:200]), default=0)
            if len(raw_lines) <= 3 and longest > 1200:
                skipped['minified'].append(rel)
                continue
            kept.append({'path': rel, 'full': full, 'text': text,
                         'lines': raw_lines, 'size': size})
    return kept, skipped


def score_file(rel):
    """核心业务代码优先：审查员希望前 30 页看到业务逻辑，而不是配置和测试。"""
    low = rel.lower()
    s = 0
    if re.search(r'(^|/)(src|app|server|core|lib|service|domain|module|pages|components|api)/', low):
        s -= 6
    if re.search(r'(test|spec|mock|fixture|__tests__|e2e)', low):
        s += 8
    if re.search(r'(config|setting|constant|\.d\.ts$|types?\.)', low):
        s += 4
    if re.search(r'(util|helper|common|base)', low):
        s += 2
    if re.search(r'(migration|schema|seed|locale|i18n|\.json$|\.yml$|\.yaml$|\.ini$|\.toml$)', low):
        s += 6
    if re.search(r'(index|main|entry|\.ets$)', low):
        s -= 2
    s += low.count('/') * 0.1
    return (s, len(low), low)


def build_document_lines(files, max_units, with_file_headers=True):
    """把源码拼成「排版行」序列：去空行、超宽换行、按文件加分隔行。"""
    out = []
    total = len(files)
    for i, f in enumerate(files, 1):
        if with_file_headers:
            head = '===== 文件 %d/%d  %s =====' % (i, total, f['path'])
            out.extend(wrap_line(head, max_units))
        for raw in f['lines']:
            line = normalize_line(raw)
            if not line.strip():
                continue
            out.extend(wrap_line(line, max_units))
    return out


def wrap_line(line, max_units):
    if display_width(line) <= max_units:
        return [line]
    chunks, buf, w = [], [], 0
    for ch in line:
        cw = 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1
        if w + cw > max_units:
            chunks.append(''.join(buf))
            buf, w = [ch], cw
        else:
            buf.append(ch)
            w += cw
    if buf:
        chunks.append(''.join(buf))
    return chunks


def paginate(lines, per_page):
    pages = []
    for i in range(0, len(lines), per_page):
        pages.append(lines[i:i + per_page])
    return pages


def emit_pages(doc, pages, line_pt, font_size, font_ascii='Consolas', start_break=True):
    """逐页写入；每页首段使用 page_break_before 强制分页，保证行数精确。"""
    first = True
    for page in pages:
        for idx, line in enumerate(page):
            para = doc.add_paragraph()
            pf = para.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = Pt(line_pt)
            pf.widow_control = False
            pf.alignment = 0
            if idx == 0 and not (first and not start_break):
                pf.page_break_before = True
            run = para.add_run(line if line else ' ')
            set_run_font(run, font_ascii, '宋体', font_size)
        first = False


def make_doc(title, font_size, total_pages_literal, font_ascii='Consolas'):
    doc = new_document('宋体', font_ascii, font_size)
    section = doc.sections[0]
    setup_page(section)
    build_header(section, title, total_pages=total_pages_literal,
                 font_ascii=font_ascii)
    set_update_fields(doc)
    return doc, section


def main():
    ap = argparse.ArgumentParser(description='生成软著程序鉴别材料（源代码 Word）')
    ap.add_argument('--root', default=None, help='项目根目录（缺省取 00_基本信息.json）')
    ap.add_argument('--name', default=None, help='软件全称（缺省取 00_基本信息.json）')
    ap.add_argument('--version', default=None, help='版本号（缺省取配置，默认 V1.0）')
    ap.add_argument('--outdir', default=None, help='输出目录（缺省取配置，默认 ./软著材料）')
    ap.add_argument('--config', default=None, help='配置文件路径 00_基本信息.json')
    ap.add_argument('--font-ascii', default='Consolas', help='西文/代码字体，缺字体时可换 Courier New')
    ap.add_argument('--lines-per-page', type=int, default=50, help='每页行数，默认 50')
    ap.add_argument('--excerpt-pages', type=int, default=30, help='前后各取页数，默认 30')
    ap.add_argument('--line-pt', type=float, default=14.0, help='固定行距（磅），默认 14')
    ap.add_argument('--font-size', type=float, default=9.0, help='字号（磅），默认 9')
    ap.add_argument('--max-units', type=int, default=96, help='单行最大显示宽度（半角字符数）')
    ap.add_argument('--number-mode', choices=['continue', 'original'], default='continue',
                    help='第二份页码：continue=1..60 连续（默认）；original=沿用完整稿真实页码')
    ap.add_argument('--no-file-headers', action='store_true', help='不加文件分隔行')
    ap.add_argument('--exclude-dir', action='append', default=[], help='追加排除目录')
    ap.add_argument('--exclude-re', default='', help='追加排除路径正则')
    args = ap.parse_args()

    cfg = cfgmod.load(args.config)
    args.name = args.name or cfgmod.get(cfg, '软件全称')
    args.version = args.version or cfgmod.get(cfg, '版本号', 'V1.0')
    args.outdir = args.outdir or cfgmod.get(cfg, '输出目录', './软著材料')
    args.root = args.root or cfgmod.get(cfg, '项目目录')
    if not args.root:
        sys.exit('缺少项目目录：加 --root <路径>，或先 python bootstrap.py --root <路径> ...')
    if not args.name:
        sys.exit('缺少软件全称：加 --name "全称"，或先 python bootstrap.py --name "全称" ...')

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        sys.exit('项目目录不存在：%s' % root)
    os.makedirs(args.outdir, exist_ok=True)

    extra_re = re.compile(args.exclude_re) if args.exclude_re else None
    files, skipped = collect_files(root, args.exclude_dir, extra_re)
    if not files:
        sys.exit('未采集到任何源代码文件，请检查 --root 与扩展名过滤')

    files.sort(key=lambda f: score_file(f['path']))
    doc_lines = build_document_lines(files, args.max_units, not args.no_file_headers)
    pages = paginate(doc_lines, args.lines_per_page)
    total_pages = len(pages)

    line_pt = args.line_pt
    # A4 29.7cm − 上 2.2cm − 下 1.8cm = 25.7cm = 728.5 磅
    capacity = (29.7 - 2.2 - 1.8) / 2.54 * 72
    needed = line_pt * args.lines_per_page
    if needed > capacity:
        sys.exit('排版不可行：%d 行 × %.1f 磅 = %.0f 磅，超出可用高度 %.0f 磅。'
                 '请减小 --line-pt 或 --lines-per-page' % (args.lines_per_page, line_pt, needed, capacity))

    title = '%s %s' % (args.name.strip(), args.version.strip())

    # ---- 第一份：完整源代码 ----
    full_doc, full_sec = make_doc(title, args.font_size, total_pages, args.font_ascii)
    emit_pages(full_doc, pages, line_pt, args.font_size, args.font_ascii)
    full_path = os.path.join(args.outdir, '源代码_完整.docx')
    full_doc.save(full_path)

    # ---- 第二份：前 N 页 + 后 N 页（仅当超过 2N 页）----
    excerpt_path = None
    if total_pages > args.excerpt_pages * 2:
        head = pages[:args.excerpt_pages]
        tail = pages[-args.excerpt_pages:]
        if args.number_mode == 'continue':
            header_total = args.excerpt_pages * 2
            second_start = args.excerpt_pages + 1
        else:
            header_total = total_pages
            second_start = total_pages - args.excerpt_pages + 1

        ex_doc, ex_sec = make_doc(title, args.font_size, header_total, args.font_ascii)
        set_page_number_start(ex_sec, 1)
        emit_pages(ex_doc, head, line_pt, args.font_size, args.font_ascii)
        clone_sectpr_into_last_paragraph(ex_doc)
        set_page_number_start(ex_doc.sections[-1], second_start)
        emit_pages(ex_doc, tail, line_pt, args.font_size, args.font_ascii, start_break=False)
        excerpt_path = os.path.join(args.outdir, '源代码_前%d后%d.docx' %
                                    (args.excerpt_pages, args.excerpt_pages))
        ex_doc.save(excerpt_path)

    # ---- 统计 ----
    lang_counter = {}
    raw_lines = 0
    for f in files:
        lang = EXTS[os.path.splitext(f['path'])[1].lower()]
        lang_counter[lang] = lang_counter.get(lang, 0) + 1
        raw_lines += len([l for l in f['lines'] if l.strip()])

    stats = {
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '项目目录': root,
        '软件全称': args.name,
        '版本号': args.version,
        '页眉文字': title,
        '文件数': len(files),
        '源程序量_有效行数': raw_lines,
        '排版总行数': len(doc_lines),
        '每页行数': args.lines_per_page,
        '总页数': total_pages,
        '是否生成前后30页': bool(excerpt_path),
        '页码模式': args.number_mode,
        '语言分布': lang_counter,
        '文件清单': [{'path': f['path'], '有效行数': len([l for l in f['lines'] if l.strip()])}
                     for f in files],
        '已排除_疑似密钥文件': skipped['secret'],
        '已排除_压缩代码': skipped['minified'],
        '已排除_超大文件': skipped['large'],
    }
    stats_path = os.path.join(args.outdir, '源程序统计.json')
    with open(stats_path, 'w', encoding='utf-8') as fp:
        json.dump(stats, fp, ensure_ascii=False, indent=2)

    print(json.dumps({
        '完整稿': full_path,
        '前后30页稿': excerpt_path,
        '统计文件': stats_path,
        '源程序量_有效行数': raw_lines,
        '总页数': total_pages,
        '每页行数': args.lines_per_page,
        '文件数': len(files),
        '语言分布': lang_counter,
        '已排除疑似密钥文件': skipped['secret'],
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
