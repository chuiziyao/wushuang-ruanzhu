# -*- coding: utf-8 -*-
"""生成提交前自查单（05_自查单.md），把填表值与材料统计填进勾选项。

  python make_checklist.py --json 01_表单值.json --stats 源程序统计.json --out 05_自查单.md
"""
import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config as cfgmod  # noqa: E402

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def load(path):
    if path and os.path.isfile(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return {}


def main():
    ap = argparse.ArgumentParser(description='生成软著提交前自查单')
    ap.add_argument('--json', default='01_表单值.json', help='01_表单值.json')
    ap.add_argument('--stats', default=None, help='源程序统计.json（缺省自动查找）')
    ap.add_argument('--config', default=None, help='配置文件路径 00_基本信息.json')
    ap.add_argument('--manual', default=None, help='说明书 docx 或 md 路径（仅记录）')
    ap.add_argument('--out', default='05_自查单.md')
    args = ap.parse_args()

    args.json = args.json if os.path.isfile(args.json) else (
        cfgmod.find_up(args.json) or args.json)
    if not args.stats:
        args.stats = cfgmod.find_up('源程序统计.json')
    elif not os.path.isfile(args.stats):
        args.stats = cfgmod.find_up(args.stats) or args.stats
    cfg = cfgmod.load(args.config)
    if not os.path.isfile(args.json):
        sys.exit('找不到表单值文件：%s（先在上一级用 check_form.py 生成）' % args.json)
    form = load(args.json)
    stats = load(args.stats)

    name = form.get('软件全称', '（未填）')
    ver = form.get('版本号', '（未填）')
    header = '%s %s' % (name, ver)
    func = str(form.get('软件的主要功能', ''))
    pages = stats.get('总页数')
    lines = stats.get('源程序量_有效行数')
    per_page = stats.get('每页行数', 50)
    excerpt = stats.get('是否生成前后30页')

    L = []
    L.append('# 软著提交前自查单')
    L.append('')
    if cfgmod.get(cfg, '申请人名称'):
        L.append('- 申请人：%s' % cfg['申请人名称'])
    L.append('- 软件全称：%s' % name)
    L.append('- 版本号：%s' % ver)
    L.append('- 页眉文字：`%s`' % header)
    L.append('- 生成时间：%s' % datetime.now().strftime('%Y-%m-%d %H:%M'))
    L.append('')
    L.append('## 一、文字一致性（最高频补正原因）')
    L.append('')
    L.append('- [ ] 申请表「软件全称」＝ `%s`，与页眉、说明书封面逐字相同' % name)
    L.append('- [ ] 申请表「版本号」＝ `%s`，与页眉、说明书封面逐字相同（含是否带 V）' % ver)
    L.append('- [ ] 软件简称：%s' % ('留空' if not form.get('软件简称') else '`%s`，与全称不同' % form.get('软件简称')))
    L.append('- [ ] 说明书正文出现的模块名，与「软件的主要功能」用词一致')
    L.append('- [ ] 说明书截图中的软件名称、版本号与申请表一致')
    L.append('')
    L.append('## 二、数字口径')
    L.append('')
    L.append('- [ ] 主要功能 %d 字（要求 500–1300）' % len(func))
    L.append('- [ ] 源程序量填 %s 行，与源代码 Word 的有效行数同口径' % form.get('源程序量', '（未填）'))
    if lines:
        L.append('- [ ] 生成器统计的有效行数为 %s 行，两者应一致' % lines)
    if pages:
        L.append('- [ ] 源代码共 %s 页，每页 %s 行' % (pages, per_page))
        if excerpt:
            L.append('- [ ] 已生成「前 30 页 + 后 30 页」版本（总页数超过 60）')
        else:
            L.append('- [ ] 总页数未超过 60，提交完整稿即可')
    L.append('- [ ] 环境类五栏每栏 ≤ 50 字；编程语言其他栏 ≤ 120 字；技术特点补充 ≤ 100 字')
    L.append('')
    L.append('## 三、材料格式')
    L.append('')
    L.append('- [ ] 源代码 Word 已转 PDF，PDF 首页/末页页码与「共 N 页」对得上')
    L.append('- [ ] Word 导出 PDF，未使用缩放打印')
    L.append('- [ ] 说明书 Word 已转 PDF；超过 60 页的按第 1–30 页与最后 30 页提取')
    L.append('- [ ] 每个 PDF ≤ 20MB，文件名含软件全称')
    L.append('- [ ] 源代码已脱敏：无 API 密钥、数据库口令、内网 IP、真实用户数据')
    L.append('- [ ] 源代码未包含 node_modules / vendor / 构建产物 / 压缩代码')
    L.append('')
    L.append('## 四、系统填报')
    L.append('')
    L.append('- [ ] 企业实名认证已通过，主体名称与营业执照全称一致')
    L.append('- [ ] 开发完成日期 ≥ 主体成立日期，且 ≤ 今天')
    L.append('- [ ] 发表状态为已发表的，首次发表日期 ≥ 开发完成日期，地点已填')
    L.append('- [ ] 「确认信息」页逐字复核过一遍，再点「填报完成」')
    L.append('')
    L.append('## 五、签章页（缺一步就不予受理）')
    L.append('')
    L.append('- [ ] 在线打印确认签章页，未改内容、未改比例')
    L.append('- [ ] 企业公章鲜章清晰，多页的已骑缝或逐页盖章')
    L.append('- [ ] 已扫描为 PDF 并回传系统')
    L.append('- [ ] 回传在系统规定期限内完成，未逾期')
    L.append('')
    L.append('## 六、提交之后')
    L.append('')
    L.append('- [ ] 记录受理号与提交时间')
    L.append('- [ ] 收到补正通知后 30 日内处理')
    L.append('- [ ] 下证后下载电子证书并归档')
    L.append('')
    L.append('> 本单由脚本按填表值与材料统计生成，勾选动作由申请人完成。')

    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    print('已生成：%s' % os.path.abspath(args.out))


if __name__ == '__main__':
    main()
