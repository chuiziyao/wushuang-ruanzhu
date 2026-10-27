# -*- coding: utf-8 -*-
"""扫描项目，产出填表所需的「开发事实」JSON，供 AI 起草申请表字段。

覆盖：语言与行数分布、依赖清单、框架/技术栈识别、目录结构、关键文件、
git 提交区间（开发完成日期参考）、疑似密钥文件（必须脱敏排除）。

用法：
  python scan_project.py --root D:/proj --out 02_项目扫描.json
"""
import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

LANG_EXT = {
    '.py': 'Python', '.js': 'JavaScript', '.mjs': 'JavaScript', '.cjs': 'JavaScript',
    '.ts': 'TypeScript', '.tsx': 'TypeScript', '.jsx': 'JavaScript',
    '.vue': 'Vue', '.kt': 'Kotlin', '.java': 'Java', '.go': 'Go',
    '.c': 'C', '.h': 'C', '.cpp': 'C++', '.cc': 'C++', '.hpp': 'C++',
    '.cs': 'C#', '.swift': 'Swift', '.rs': 'Rust', '.php': 'PHP',
    '.rb': 'Ruby', '.m': 'Objective-C', '.mm': 'Objective-C', '.dart': 'Dart',
    '.lua': 'Lua', '.ets': 'ArkTS', '.sql': 'SQL', '.html': 'HTML',
    '.htm': 'HTML', '.css': 'CSS', '.scss': 'SCSS', '.less': 'LESS',
    '.sh': 'Shell', '.ps1': 'PowerShell', '.asm': 'Assembly', '.s': 'Assembly',
}

SKIP_DIRS = {
    'node_modules', '.git', 'dist', 'build', 'out', 'target', '.gradle', '.idea',
    '.vscode', '__pycache__', 'venv', '.venv', 'env', 'site-packages', 'vendor',
    'coverage', '.next', '.nuxt', '.workbuddy', 'Library', 'Temp', 'Logs',
    'obj', 'bin', 'Pods', 'DerivedData', '.dart_tool', 'release', 'debug',
}

SECRET_RE = re.compile(
    r'(secret|credential|passwd|password|\.env|id_rsa|\.pem$|\.pfx$|token|'
    r'apikey|api_key|private_key)', re.I)

FRAMEWORK_HINTS = [
    ('package.json', 'Node.js / npm'),
    ('requirements.txt', 'Python / pip'),
    ('pyproject.toml', 'Python / PEP 518'),
    ('pom.xml', 'Java / Maven'),
    ('build.gradle', 'Java-Kotlin / Gradle'),
    ('pubspec.yaml', 'Dart / Flutter'),
    ('oh-package.json5', 'HarmonyOS / ArkTS'),
    ('Cargo.toml', 'Rust / Cargo'),
    ('go.mod', 'Go Modules'),
    ('Podfile', 'iOS / CocoaPods'),
    ('CMakeLists.txt', 'C/C++ / CMake'),
]

CODE_SIGNS = [
    ('fastapi', 'FastAPI（Python Web 框架）'),
    ('flask', 'Flask（Python Web 框架）'),
    ('django', 'Django（Python Web 框架）'),
    ('react', 'React（前端框架）'),
    ('vue', 'Vue（前端框架）'),
    ('@angular', 'Angular（前端框架）'),
    ('express', 'Express（Node 服务端框架）'),
    ('electron', 'Electron（桌面应用）'),
    ('unityengine', 'Unity（游戏引擎）'),
    ('jetpack compose', 'Jetpack Compose（Android UI）'),
    ('@ohos', 'HarmonyOS ArkTS'),
    ('playwright', 'Playwright（浏览器自动化）'),
    ('sqlalchemy', 'SQLAlchemy（ORM）'),
    ('torch', 'PyTorch（深度学习）'),
    ('tensorflow', 'TensorFlow（深度学习）'),
    ('openai', 'OpenAI SDK'),
    ('langchain', 'LangChain'),
    ('sqlite', 'SQLite'),
    ('mysql', 'MySQL'),
    ('redis', 'Redis'),
    ('nginx', 'Nginx'),
]


def read_json(path):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def walk_files(root):
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                       if d not in SKIP_DIRS and not d.startswith('.')]
        for fn in filenames:
            out.append(os.path.join(dirpath, fn))
    return out


def git_info(root):
    info = {}
    try:
        first = subprocess.run(['git', '-C', root, 'log', '--reverse',
                                '--format=%ad', '--date=short'],
                               capture_output=True, text=True, timeout=20)
        last = subprocess.run(['git', '-C', root, 'log', '-1', '--format=%ad',
                               '--date=short'],
                              capture_output=True, text=True, timeout=20)
        cnt = subprocess.run(['git', '-C', root, 'rev-list', '--count', 'HEAD'],
                             capture_output=True, text=True, timeout=20)
        dates = [d for d in first.stdout.split('\n') if d.strip()]
        if dates:
            info['首次提交'] = dates[0]
            info['最近提交'] = last.stdout.strip()
            info['提交次数'] = cnt.stdout.strip()
    except Exception:
        pass
    return info


def main():
    ap = argparse.ArgumentParser(description='扫描项目，输出软件著作权申请所需开发事实')
    ap.add_argument('--root', required=True)
    ap.add_argument('--out', default='02_项目扫描.json')
    ap.add_argument('--top', type=int, default=30, help='关键文件返回数量')
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        sys.exit('目录不存在：%s' % root)

    files = walk_files(root)
    lang_stat = {}
    total_files = 0
    total_lines = 0
    total_nonblank = 0
    total_bytes = 0
    secrets = []
    big_files = []

    for full in files:
        rel = os.path.relpath(full, root).replace('\\', '/')
        ext = os.path.splitext(full)[1].lower()
        try:
            size = os.path.getsize(full)
        except OSError:
            continue
        if SECRET_RE.search(rel) and size > 0:
            secrets.append(rel)
        total_files += 1
        total_bytes += size
        if ext not in LANG_EXT or size > 2 * 1024 * 1024:
            continue
        try:
            with open(full, encoding='utf-8', errors='ignore') as f:
                text = f.read()
        except OSError:
            continue
        lines = text.split('\n')
        nonblank = sum(1 for l in lines if l.strip())
        lang = LANG_EXT[ext]
        st = lang_stat.setdefault(lang, {'文件数': 0, '总行数': 0, '有效行数': 0})
        st['文件数'] += 1
        st['总行数'] += len(lines)
        st['有效行数'] += nonblank
        total_lines += len(lines)
        total_nonblank += nonblank
        big_files.append((nonblank, rel, lang))

    # 依赖与框架
    manifests = []
    for name, desc in FRAMEWORK_HINTS:
        for full in files:
            if os.path.basename(full) == name:
                manifests.append({'文件': os.path.relpath(full, root).replace('\\', '/'),
                                  '类型': desc})
    deps = []
    pkg = None
    for full in files:
        if os.path.basename(full) == 'package.json' and 'node_modules' not in full:
            pkg = read_json(full)
            if pkg:
                deps += list((pkg.get('dependencies') or {}).keys())
                deps += list((pkg.get('devDependencies') or {}).keys())
            break
    req = None
    for full in files:
        if os.path.basename(full) == 'requirements.txt':
            try:
                with open(full, encoding='utf-8', errors='ignore') as f:
                    req = [l.strip() for l in f if l.strip()
                           and not l.strip().startswith('#')]
            except OSError:
                pass
            break

    blob = ''
    sample = 0
    for full in files:
        ext = os.path.splitext(full)[1].lower()
        if ext in ('.py', '.js', '.ts', '.vue', '.kt', '.java', '.go', '.dart', '.ets'):
            try:
                with open(full, encoding='utf-8', errors='ignore') as f:
                    blob += f.read(4000).lower()
            except OSError:
                continue
            sample += 1
            if sample > 60:
                break

    frameworks = []
    haystack = blob + ' ' + ' '.join(d.lower() for d in deps) + ' ' + ' '.join(
        (req or [])).lower()
    for sign, label in CODE_SIGNS:
        if sign in haystack:
            frameworks.append(label)

    dirs = sorted({os.path.relpath(d, root).replace('\\', '/')
                   for d in [os.path.join(dp) for dp, dn, fn in os.walk(root)
                             if dp != root]})[:40]

    big_files.sort(reverse=True)
    key_files = [{'路径': p, '有效行数': n, '语言': l} for n, p, l in big_files[:args.top]]

    langs_sorted = sorted(lang_stat.items(), key=lambda kv: -kv[1]['有效行数'])

    result = {
        '扫描时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '项目目录': root,
        '项目名': (pkg or {}).get('name') or os.path.basename(root),
        '文件总数': total_files,
        '总体积MB': round(total_bytes / 1024 / 1024, 2),
        '语言分布': {k: v for k, v in langs_sorted},
        '代码总行数': total_lines,
        '代码有效行数': total_nonblank,
        '清单文件': manifests,
        '主要依赖': (deps[:40] + (req or [])[:40]),
        '识别技术栈': frameworks,
        '目录结构': dirs,
        '关键文件': key_files,
        'git': git_info(root),
        '疑似密钥文件_必须排除或脱敏': secrets[:40],
        '建议填表值': {
            '编程语言': [k for k, _ in langs_sorted[:6]],
            '源程序量': total_nonblank,
            '源程序量说明': '有效行数（去空行）；若按含空行口径统计为 %d 行' % total_lines,
            '开发完成日期参考': (git_info(root) or {}).get('最近提交', '以实际固化为准'),
        },
        '提示': [
            '源程序量建议取「有效行数」并在填表后与源代码 Word 的总行数口径保持一致',
            '疑似密钥文件已列出，生成源代码前必须排除或替换为占位符',
            '技术栈识别为关键词匹配结果，需开发者确认后再写入申请表',
        ],
    }
    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(json.dumps({
        '输出': os.path.abspath(args.out),
        '语言分布': {k: v['有效行数'] for k, v in langs_sorted},
        '代码有效行数': total_nonblank,
        '识别技术栈': frameworks,
        '疑似密钥文件数': len(secrets),
        'git': result['git'],
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
