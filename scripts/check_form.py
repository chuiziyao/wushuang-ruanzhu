# -*- coding: utf-8 -*-
"""校验申请表字段草稿，并输出可逐项复制的《填报清单》。

输入：01_表单值.json（AI 起草，字段名与下表一致）
输出：01_填报清单.md（按官网六个页签分组；长字段单独给复制块；含错误与警告）

  python check_form.py --json 01_表单值.json --out 01_填报清单.md \
      --scan 02_项目扫描.json --founded 2016-03-18
  python check_form.py --json 01_表单值.json --init      # 生成空白模板
"""
import argparse
import json
import os
import re
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as cfgmod  # noqa: E402

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

# (字段, 页签, 类型, 限制, 必填, 说明)
SINGLE_OPTIONS = {
    '权利取得方式': ['原始取得', '继受取得'],
    '权利范围': ['全部权利', '部分权利'],
    '软件说明': ['原创', '修改（含翻译软件、合成软件）'],
    '开发方式': ['单独开发', '合作开发', '委托开发', '下达任务开发'],
    '发表状态': ['未发表', '已发表'],
    '程序鉴别材料': ['一般交存', '例外交存'],
    '文档鉴别材料': ['一般交存', '例外交存'],
    '著作权人国别': ['中国'],
}

TECH_TAGS = [
    'APP', '游戏软件', '教育软件', '金融软件', '医疗软件', '地理信息软件',
    '云计算软件', '信息安全软件', '大数据软件', '人工智能软件', 'VR软件',
    '5G软件', '小程序', '物联网软件', '智慧城市软件',
]

LANG_TAGS = [
    'Assembly language', 'C', 'C#', 'C++', 'Delphi/Object Pascal', 'Go', 'HTML',
    'Java', 'JavaScript', 'Kotlin', 'Python', 'PHP', 'Ruby', 'SQL', 'Swift',
    'TypeScript', 'Visual Basic', '其他',
]

LIMITS = {
    '开发的硬件环境': (None, 50),
    '运行的硬件环境': (None, 50),
    '开发该软件的操作系统': (None, 50),
    '软件开发环境/开发工具': (None, 50),
    '该软件的运行平台/操作系统': (None, 50),
    '软件运行支撑环境/支持软件': (None, 50),
    '其他编程语言': (None, 120),
    '开发目的': (None, 50),
    '面向领域/行业': (None, 50),
    '软件的主要功能': (500, 1300),
    '技术特点补充': (None, 100),
    '软件全称': (None, 40),
    '软件简称': (None, 15),
}

TEMPLATE = {
    '权利取得方式': '原始取得',
    '软件全称': '（品牌/主体 + 核心功能 + 软件/系统/平台，不含版本号）',
    '软件简称': '',
    '版本号': 'V1.0',
    '权利范围': '全部权利',
    '软件分类': '应用软件',
    '软件说明': '原创',
    '开发方式': '单独开发',
    '开发完成日期': 'YYYY-MM-DD',
    '发表状态': '未发表',
    '首次发表日期': '',
    '首次发表地点': '',
    '著作权人国别': '中国',
    '著作权人省市': '江苏 / 南京',
    '著作权人类别': '企业法人',
    '著作权人名称': '（营业执照全称，逐字一致）',
    '证件类型': '统一社会信用代码证书',
    '证件号码': '（18 位，逐字核对）',
    '开发的硬件环境': '',
    '运行的硬件环境': '',
    '开发该软件的操作系统': '',
    '软件开发环境/开发工具': '',
    '该软件的运行平台/操作系统': '',
    '软件运行支撑环境/支持软件': '',
    '编程语言': [],
    '其他编程语言': '',
    '源程序量': 0,
    '开发目的': '',
    '面向领域/行业': '',
    '软件的主要功能': '',
    '软件的技术特点': [],
    '技术特点补充': '',
    '程序鉴别材料': '一般交存',
    '文档鉴别材料': '一般交存',
    '文档类型': '操作说明书',
    '其他相关证明文件': '',
}

PAGE_ORDER = [
    ('① 软件申请信息', ['权利取得方式', '软件全称', '软件简称', '版本号', '权利范围']),
    ('② 软件开发信息', ['软件分类', '软件说明', '开发方式', '开发完成日期', '发表状态',
                   '首次发表日期', '首次发表地点', '著作权人国别', '著作权人省市',
                   '著作权人类别', '著作权人名称', '证件类型', '证件号码']),
    ('③ 软件功能与特点（一）：环境与语言', [
        '开发的硬件环境', '运行的硬件环境', '开发该软件的操作系统',
        '软件开发环境/开发工具', '该软件的运行平台/操作系统',
        '软件运行支撑环境/支持软件', '编程语言', '其他编程语言', '源程序量']),
    ('③ 软件功能与特点（二）：功能描述', [
        '开发目的', '面向领域/行业', '软件的主要功能', '软件的技术特点', '技术特点补充']),
    ('③ 软件功能与特点（三）：鉴别材料', [
        '程序鉴别材料', '文档鉴别材料', '文档类型', '其他相关证明文件']),
]


def clen(s):
    return len(str(s).strip()) if s is not None else 0


def check(data, founded=None, scan=None):
    errors, warns, infos = [], [], []
    today = date.today()

    # 单选项
    for key, opts in SINGLE_OPTIONS.items():
        val = data.get(key)
        if val in (None, ''):
            if key in ('著作权人国别',):
                continue
            errors.append('「%s」未填写，可选：%s' % (key, ' / '.join(opts)))
        elif val not in opts:
            errors.append('「%s」取值「%s」不在允许项内：%s' % (key, val, ' / '.join(opts)))

    # 软件全称
    name = (data.get('软件全称') or '').strip()
    if not name:
        errors.append('「软件全称」未填写')
    else:
        if re.search(r'[Vv]?\d+\.\d+', name):
            errors.append('「软件全称」含版本号，版本号应单独填在「版本号」栏')
        if re.search(r'\s', name):
            warns.append('「软件全称」含空格，确认与页眉、手册三处逐字一致')
        if clen(name) > 40:
            errors.append('「软件全称」%d 字，超过 40 字上限' % clen(name))
        elif clen(name) > 25:
            warns.append('「软件全称」%d 字，偏长，建议控制在 25 字以内' % clen(name))
        if not re.search(r'(软件|系统|平台|程序|工具|引擎)$', name):
            warns.append('「软件全称」建议以「软件 / 系统 / 平台」结尾，与审查习惯一致')

    # 简称
    short = (data.get('软件简称') or '').strip()
    if short in ('无', '無', '/', '-', '—'):
        errors.append('「软件简称」无简称时应留空，不得填写「无」或符号')
    elif short and short == name:
        errors.append('「软件简称」不得与全称相同')
    if short and clen(short) > 15:
        errors.append('「软件简称」%d 字，超过 15 字上限' % clen(short))

    # 版本号
    ver = (data.get('版本号') or '').strip()
    if not ver:
        errors.append('「版本号」未填写')
    elif not re.match(r'^[Vv]?\d+(\.\d+){0,2}$', ver):
        errors.append('「版本号」格式不规范，应为 V1.0 或 1.0 形式')
    elif not ver.upper().startswith('V'):
        warns.append('「版本号」未带 V 前缀（%s）。可与官网格式一致，但页眉必须逐字相同' % ver)

    # 日期逻辑
    def parse_d(s):
        if not s:
            return None
        s = str(s).strip().replace('/', '-').replace('.', '-')
        s = re.sub(r'[年月]', '-', s).replace('日', '')
        try:
            return datetime.strptime(s, '%Y-%m-%d').date()
        except ValueError:
            return None

    done = parse_d(data.get('开发完成日期'))
    if not done:
        errors.append('「开发完成日期」缺失或格式错误，应为 YYYY-MM-DD')
    else:
        if done > today:
            errors.append('「开发完成日期」%s 晚于今天，不合法' % done)
        if founded and done < founded:
            errors.append('「开发完成日期」%s 早于著作权人成立日期 %s，必被要求补正'
                          % (done, founded))
    if data.get('发表状态') == '已发表':
        pub = parse_d(data.get('首次发表日期'))
        if not pub:
            errors.append('发表状态为「已发表」，必须填写「首次发表日期」')
        else:
            if pub > today:
                errors.append('「首次发表日期」晚于今天')
            if done and pub < done:
                errors.append('「首次发表日期」早于开发完成日期')

    # 字数限制
    for key, (lo, hi) in LIMITS.items():
        val = data.get(key)
        if val in (None, ''):
            if key != '软件简称':
                warns.append('「%s」为空（限 %s 字）' % (key, hi))
            continue
        n = clen(val)
        if lo and n < lo:
            errors.append('「%s」%d 字，低于下限 %d 字（该栏要求 %d–%d 字）' % (key, n, lo, lo, hi))
        if hi and n > hi:
            errors.append('「%s」%d 字，超过上限 %d 字' % (key, n, hi))

    # 技术特点 / 编程语言
    tags = data.get('软件的技术特点') or []
    if not isinstance(tags, list):
        errors.append('「软件的技术特点」应为数组')
    else:
        bad = [t for t in tags if t not in TECH_TAGS]
        if bad:
            warns.append('「软件的技术特点」含非官网标签：%s' % '、'.join(bad))
        if not tags:
            errors.append('「软件的技术特点」至少选择一个标签')

    langs = data.get('编程语言') or []
    if not isinstance(langs, list):
        errors.append('「编程语言」应为数组')
    elif not langs:
        errors.append('「编程语言」至少选择一项（不在列表的写入「其他编程语言」）')
    else:
        bad = [t for t in langs if t not in LANG_TAGS]
        if bad:
            warns.append('「编程语言」含官网标签外语言，请确认写入「其他编程语言」栏：%s'
                         % '、'.join(bad))

    # 源程序量
    lines = data.get('源程序量')
    if not lines:
        errors.append('「源程序量」未填写（应为正整数，单位：行）')
    else:
        try:
            lines = int(str(lines).replace(',', '').strip())
        except ValueError:
            errors.append('「源程序量」应为数字')
            lines = None
        if lines is not None and lines <= 0:
            errors.append('「源程序量」必须大于 0')

    # 与扫描结果交叉核对
    if scan and isinstance(scan, dict):
        stat = (scan.get('建议填表值') or {}).get('源程序量')
        if stat and lines:
            diff = abs(lines - stat) / max(stat, 1)
            if diff > 0.1:
                warns.append('「源程序量」%s 行与项目扫描的有效行数 %s 行相差 %.0f%%，'
                             '确认统计口径（是否含空行、是否剔除依赖）'
                             % (lines, stat, diff * 100))
        sug = (scan.get('建议填表值') or {}).get('编程语言') or []
        missing = [s for s in sug if s not in (langs or [])]
        if missing:
            warns.append('扫描到项目使用 %s，但申请表未勾选，请确认' % '、'.join(missing))

    if not data.get('著作权人名称'):
        errors.append('「著作权人名称」未填写')
    code = (data.get('证件号码') or '').strip()
    if code:
        if len(code) != 18:
            errors.append('「证件号码」%d 位，统一社会信用代码应为 18 位' % len(code))
        elif not re.match(r'^[0-9A-HJ-NPQRTUWXY]{18}$', code.upper()):
            warns.append('「证件号码」含疑似非法字符，请与营业执照逐字核对')

    # 三处一致性提示
    infos.append('页眉统一写法：%s %s（申请表、源代码 Word、手册 Word 三处必须逐字一致）'
                 % (name or '(软件全称)', (ver or 'V1.0')))
    infos.append('源程序量与源代码 Word 口径必须一致：Word 中的有效代码行数应等于填表行数')
    return errors, warns, infos


def render(data, errors, warns, infos, out_path):
    name = (data.get('软件全称') or '').strip()
    ver = (data.get('版本号') or '').strip()
    L = []
    L.append('# 软件著作权申请表 · 填报清单')
    L.append('')
    L.append('- 软件全称：**%s**' % (name or '（未填）'))
    L.append('- 版本号：**%s**' % (ver or '（未填）'))
    L.append('- 页眉文字（复制到 Word 页眉左侧）：`%s %s`' % (name, ver))
    L.append('- 校验结果：%s（错误 %d 项 / 提示 %d 项）'
             % ('✅ 可提交' if not errors else '❌ 需修正', len(errors), len(warns)))
    L.append('')
    if errors:
        L.append('## 必须修正')
        L.append('')
        for e in errors:
            L.append('- ❌ %s' % e)
        L.append('')
    if warns:
        L.append('## 建议核对')
        L.append('')
        for w in warns:
            L.append('- ⚠️ %s' % w)
        L.append('')
    if infos:
        L.append('## 一致性提醒')
        L.append('')
        for i in infos:
            L.append('- %s' % i)
        L.append('')

    long_fields = {'软件的主要功能', '软件全称', '面向领域/行业', '开发目的',
                   '软件开发环境/开发工具', '开发的硬件环境', '运行的硬件环境',
                   '开发该软件的操作系统', '该软件的运行平台/操作系统',
                   '软件运行支撑环境/支持软件'}

    for page_title, keys in PAGE_ORDER:
        L.append('## %s' % page_title)
        L.append('')
        L.append('| 字段 | 填写内容 | 字数 | 状态 |')
        L.append('|---|---|---|---|')
        for key in keys:
            val = data.get(key, '')
            if isinstance(val, list):
                shown = '、'.join(val) if val else '（空）'
                cnt = clen(shown)
            else:
                shown = str(val) if val not in (None, '') else '（空）'
                cnt = clen(val)
            if key in LIMITS:
                lo, hi = LIMITS[key]
                limit = ('%d–%d' % (lo, hi)) if lo else ('≤%d' % hi)
                status = '✅' if (not hi or cnt <= hi) and (not lo or cnt >= lo) else '❌'
                cnt_s = '%d / %s' % (cnt, limit)
            else:
                cnt_s = str(cnt)
                status = '✅' if shown != '（空）' else '⚠️'
            shown = shown.replace('|', '／').replace('\n', ' ')
            if len(shown) > 60:
                shown = shown[:48] + '……（全文见下方复制块）'
            L.append('| %s | %s | %s | %s |' % (key, shown, cnt_s, status))
        L.append('')

        # 长字段给独立复制块
        blocks = [k for k in keys if k in long_fields and data.get(k)]
        for k in blocks:
            L.append('**%s**（可整段复制，%d 字）' % (k, clen(data.get(k))))
            L.append('')
            L.append('```text')
            L.append(str(data.get(k)).strip())
            L.append('```')
            L.append('')

    L.append('## 提交后动作')
    L.append('')
    L.append('1. 逐项把上表内容粘进 R11 系统对应页签，注意粘贴后字数计数与本地一致。')
    L.append('2. 走到「确认信息」页逐字核对，再进入「填报完成」。')
    L.append('3. 系统生成《申请确认签章页》后打印，加盖企业公章（鲜章），扫描 PDF 上传。')
    L.append('4. 材料：程序鉴别材料 PDF（前 30 页 + 后 30 页）、文档鉴别材料 PDF、营业执照副本。')
    L.append('5. 提交后保留受理号，30 日内留意补正通知。')
    L.append('')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    return out_path


def main():
    ap = argparse.ArgumentParser(description='校验软著申请表字段草稿并生成填报清单')
    ap.add_argument('--json', default=None, help='表单值 JSON，缺省取当前或上级目录的 01_表单值.json')
    ap.add_argument('--out', default='01_填报清单.md')
    ap.add_argument('--scan', default=None, help='项目扫描 JSON（缺省自动找 02_项目扫描.json）')
    ap.add_argument('--founded', default=None, help='著作权人成立日期，缺省取配置')
    ap.add_argument('--config', default=None, help='配置文件路径 00_基本信息.json')
    ap.add_argument('--init', action='store_true', help='生成空白模板后退出')
    args = ap.parse_args()

    cfg = cfgmod.load(args.config)
    if not args.json:
        args.json = cfgmod.find_up('01_表单值.json') or '01_表单值.json'
    elif not os.path.isfile(args.json):
        args.json = cfgmod.find_up(args.json) or args.json

    if args.init:
        if os.path.exists(args.json):
            sys.exit('模板已存在，避免覆盖：%s' % args.json)
        tpl = dict(TEMPLATE)
        if cfgmod.get(cfg, '软件全称'):
            tpl['软件全称'] = cfg['软件全称']
        if cfgmod.get(cfg, '申请人名称'):
            tpl['著作权人名称'] = cfg['申请人名称']
        if cfgmod.get(cfg, '版本号'):
            tpl['版本号'] = cfg['版本号']
        if cfgmod.get(cfg, '统一社会信用代码'):
            tpl['证件号码'] = cfg['统一社会信用代码']
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(tpl, f, ensure_ascii=False, indent=2)
        print('已生成模板：%s（已按 00_基本信息.json 预填申请人、软件名、版本号）' % args.json)
        return

    with open(args.json, encoding='utf-8') as f:
        data = json.load(f)

    if not args.scan:
        args.scan = cfgmod.find_up('02_项目扫描.json')
    if not args.founded:
        args.founded = cfgmod.get(cfg, '成立日期')
    scan = None
    if args.scan and os.path.isfile(args.scan):
        with open(args.scan, encoding='utf-8') as f:
            scan = json.load(f)

    founded = None
    if args.founded:
        try:
            founded = datetime.strptime(args.founded, '%Y-%m-%d').date()
        except ValueError:
            sys.exit('--founded 格式应为 YYYY-MM-DD')

    errors, warns, infos = check(data, founded, scan)
    render(data, errors, warns, infos, args.out)
    print(json.dumps({
        '清单': os.path.abspath(args.out),
        '错误': errors,
        '提示': warns,
    }, ensure_ascii=False, indent=2))
    sys.exit(0 if not errors else 1)


if __name__ == '__main__':
    main()
