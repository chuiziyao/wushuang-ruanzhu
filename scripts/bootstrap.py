# -*- coding: utf-8 -*-
"""首次使用自检：查环境、装缺件、记下申请人/软件名/版本号。

设计原则（针对国内网络）：
  · 能不下载就不下载。PDF 出稿走本机 Word/WPS 的 COM，零依赖。
  · 必须下载的只有两个纯 Python 小包（python-docx、pypdf），走清华镜像，几秒完成。
  · 装不上不阻塞：会降级并在报告里写清下一步怎么做。

用法：
  python bootstrap.py                                  # 只自检 + 提示待填三项
  python bootstrap.py --owner "某某公司" --name "某某软件" --version V1.0
  python bootstrap.py --owner ... --name ... --install  # 顺手补装缺失依赖
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config  # noqa: E402

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

MIRRORS = [
    'https://pypi.tuna.tsinghua.edu.cn/simple',
    'https://mirrors.aliyun.com/pypi/simple',
]

PS_PROBE = (
    "$r=@();"
    "foreach($p in @('Word.Application','KWPS.Application','WPS.Application')){"
    "try{$t=[Type]::GetTypeFromProgID($p); if($t){$r+=$p}}catch{}};"
    "'PROGID:' + ($r -join ',');"
    "'SOFFICE:' + (Test-Path 'C:\\Program Files\\LibreOffice\\program\\soffice.exe')"
)


def pad(text, width):
    """按显示宽度补齐：中文算 2 格，避免列错位。"""
    import unicodedata
    w = sum(2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1 for ch in text)
    return text + ' ' * max(0, width - w)


def check_module(name):
    try:
        mod = __import__(name)
        ver = getattr(mod, '__version__', '')
        return True, str(ver)
    except Exception:
        return False, ''


def probe_office():
    """用 PowerShell 查 COM ProgID，只输出 ASCII，避免中文编码问题。"""
    info = {'progids': [], 'soffice': False}
    pw = shutil.which('powershell') or shutil.which('pwsh')
    if pw:
        try:
            out = subprocess.run([pw, '-NoProfile', '-NonInteractive', '-Command', PS_PROBE],
                                 capture_output=True, text=True, timeout=45)
            for line in (out.stdout or '').splitlines():
                line = line.strip()
                if line.startswith('PROGID:'):
                    info['progids'] = [x for x in line[7:].split(',') if x.strip()]
                elif line.startswith('SOFFICE:'):
                    info['soffice'] = line[8:].strip().lower() == 'true'
        except Exception:
            pass
    for cand in ('C:/Program Files/LibreOffice/program/soffice.exe',
                 'C:/Program Files (x86)/LibreOffice/program/soffice.exe'):
        if os.path.isfile(cand):
            info['soffice'] = True
    return info


def check_fonts():
    win_fonts = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
    targets = {'Consolas': ['consola.ttf'], '宋体': ['simsun.ttc'], '黑体': ['simhei.ttf']}
    found = {}
    for name, files in targets.items():
        found[name] = any(os.path.isfile(os.path.join(win_fonts, f)) for f in files)
    return found


def pip_install(pkg):
    for mirror in MIRRORS:
        for args in (['--quiet'], []):
            cmd = [sys.executable, '-m', 'pip', 'install', pkg, '-i', mirror,
                   '--disable-pip-version-check', '--no-input'] + args
            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
                if r.returncode == 0:
                    return True, mirror
            except Exception:
                continue
    return False, ''


def main():
    ap = argparse.ArgumentParser(description='软著材料环境自检与基本信息登记')
    ap.add_argument('--owner', help='申请人（著作权人）营业执照全称')
    ap.add_argument('--name', help='软件全称')
    ap.add_argument('--version', default=None, help='版本号，默认 V1.0')
    ap.add_argument('--credit-code', default=None, help='统一社会信用代码')
    ap.add_argument('--founded', default=None, help='申请人成立日期 YYYY-MM-DD')
    ap.add_argument('--root', default=None, help='待申请软件的项目目录')
    ap.add_argument('--outdir', default=None, help='材料输出目录')
    ap.add_argument('--config', default=None, help='配置文件路径')
    ap.add_argument('--install', action='store_true', help='自动补装缺失依赖（清华镜像）')
    ap.add_argument('--no-install', action='store_true', help='只报告，不安装')
    ap.add_argument('--json', action='store_true', help='只输出 JSON')
    args = ap.parse_args()

    rows = []
    ok_all = True

    py_ok = sys.version_info >= (3, 9)
    rows.append(('Python', '%d.%d.%d' % sys.version_info[:3], py_ok,
                 '' if py_ok else '需要 3.9 及以上'))

    mods = {}
    for mod, label, required in (('docx', 'python-docx（生成 Word）', True),
                                 ('pypdf', 'pypdf（合并 PDF）', False)):
        ok, ver = check_module(mod)
        mods[mod] = {'ok': ok, 'ver': ver, 'label': label, 'required': required}
        rows.append((label, ver or '未安装', ok,
                     '' if ok else ('必需' if required else '可选，说明书超 60 页时才需要')))

    office = probe_office()
    has_word = 'Word.Application' in office['progids']
    has_wps = any(p.startswith(('KWPS', 'WPS')) for p in office['progids'])
    backend = 'Word' if has_word else ('WPS' if has_wps else ('LibreOffice' if office['soffice'] else ''))
    rows.append(('Microsoft Word', '已安装' if has_word else '未检测到', has_word,
                 '' if has_word else '可用 WPS 代替'))
    rows.append(('WPS Office', '已安装' if has_wps else '未检测到', has_wps,
                 '' if has_wps else ''))
    rows.append(('LibreOffice', '已安装' if office['soffice'] else '未检测到',
                 office['soffice'], '备选出稿方式，没有也行'))
    rows.append(('PDF 直接出稿', backend or '不可用', bool(backend),
                 '' if backend else '只出 docx，需用户在 Word 里手动另存 PDF'))

    fonts = check_fonts()
    rows.append(('等宽字体 Consolas', '已安装' if fonts['Consolas'] else '缺失',
                 fonts['Consolas'],
                 '' if fonts['Consolas'] else '可改用 --font-ascii "Courier New"'))
    rows.append(('中文字体 宋体', '已安装' if fonts['宋体'] else '缺失', fonts['宋体'], ''))

    # 补装缺失依赖
    installed = []
    need = [m for m, v in mods.items() if not v['ok']]
    if need and args.install and not args.no_install:
        for mod in need:
            pkg = 'python-docx' if mod == 'docx' else 'pypdf'
            ok, mirror = pip_install(pkg)
            if ok:
                mods[mod]['ok'] = True
                installed.append('%s（%s）' % (pkg, '清华镜像' if 'tuna' in mirror else '阿里镜像'))
    for mod, v in mods.items():
        if v['required'] and not v['ok']:
            ok_all = False

    # 写配置
    cfg = config.load(args.config)
    cfg_path = None
    if any([args.owner, args.name, args.version, args.credit_code,
            args.founded, args.root, args.outdir]):
        for k, v in (('申请人名称', args.owner), ('软件全称', args.name),
                     ('版本号', args.version), ('统一社会信用代码', args.credit_code),
                     ('成立日期', args.founded), ('项目目录', args.root),
                     ('输出目录', args.outdir)):
            if v:
                cfg[k] = v
        cfg.setdefault('版本号', 'V1.0')
        cfg_path = config.save(cfg, args.config)

    missing = [k for k in ('申请人名称', '软件全称') if not cfg.get(k)]
    if not cfg.get('版本号'):
        missing.append('版本号')

    report = {
        '环境': [{'项目': r[0], '结果': r[1], '状态': 'OK' if r[2] else 'NG',
                '说明': r[3]} for r in rows],
        '必需项缺失': [r[0] for r in rows if not r[2]
                   and r[0] in {'Python', 'python-docx（生成 Word）', 'PDF 直接出稿'}],
        '已安装': installed,
        '可生成Word': mods['docx']['ok'],
        'PDF出稿方式': backend or '无（需手动另存）',
        '可合并PDF': mods['pypdf']['ok'],
        '配置文件': cfg_path,
        '待填信息': missing,
        '结论': '可用' if ok_all else 'python-docx 缺失，无法生成 Word',
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
        sys.exit(0 if ok_all else 2)

    must = {'Python', 'python-docx（生成 Word）', 'PDF 直接出稿', '中文字体 宋体'}
    print('== 环境自检 ==')
    for name, res, ok, note in rows:
        if ok:
            status = 'OK'
        else:
            status = 'NG' if name in must else '—'
        print('  %s %s %s%s' % (pad(name, 26), pad(res, 12), status,
                                ('  ← ' + note) if note and not ok else ''))
    if installed:
        print('  已自动安装：%s' % '、'.join(installed))
    print('')
    print('== 出稿能力 ==')
    print('  Word 文档      %s' % ('可生成' if mods['docx']['ok'] else '不可用'))
    print('  PDF 出稿       %s' % ('用 %s 直接导出（不下载任何库）' % backend
                                  if backend else '无，只能出 docx'))
    print('  PDF 分页合并   %s' % ('可' if mods['pypdf']['ok'] else '不可（说明书超 60 页时需手动截页）'))
    print('')
    if cfg_path:
        print('== 已登记信息 ==  %s' % cfg_path)
        for k in ('申请人名称', '软件全称', '版本号', '统一社会信用代码',
                  '成立日期', '项目目录', '输出目录'):
            if cfg.get(k):
                print('  %-12s %s' % (k, cfg[k]))
        print('')
    if missing:
        print('== 还要问用户三项 ==')
        for k in missing:
            print('  %s' % k)
        print('')
        print('  登记命令：python bootstrap.py --owner "申请人全称" --name "软件全称" --version V1.0')
        print('')
    if not ok_all:
        print('!! python-docx 缺失：请执行 python bootstrap.py --install')
        sys.exit(2)

    print('== 下一步 ==')
    print('  1) python scan_project.py --root <项目目录> --out 02_项目扫描.json')
    print('  2) python check_form.py --json 01_表单值.json --init   # 已按登记信息预填')
    print('  3) python gen_source_docx.py --outdir <输出目录>')
    print('  4) python gen_manual_docx.py --md 操作说明书.md')
    print('  5) python docx2pdf.py --input 源代码_完整.docx --output 源代码_完整.pdf')
    print('')
    print('遇到问题：QQ 群 170621740（作者 锤无双）')


if __name__ == '__main__':
    main()
