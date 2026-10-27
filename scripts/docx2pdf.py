# -*- coding: utf-8 -*-
"""把 docx 导成 PDF（走本机 Word / WPS / LibreOffice，零下载），并支持按页截取与合并。

为什么不依赖第三方 PDF 转换库：
  · PyMuPDF 之类的二进制包体积大、国内拉取慢，且多数机器并不需要；
  · 用户机器上本来就有 Word/WPS，COM 调用是现成的，出稿质量也最好（字体、页眉页码与 Word 一致）。

用法：
  python docx2pdf.py --input 源代码_完整.docx --output x.pdf            # 整本
  python docx2pdf.py --input 说明书.docx --output head.pdf --pages 1-30
  python docx2pdf.py --input 说明书.docx --output tail.pdf --last 30
  python docx2pdf.py --merge out.pdf head.pdf tail.pdf
  python docx2pdf.py --input a.docx --info                              # 只报页数
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$job = Get-Content -LiteralPath '__JOB__' -Encoding UTF8 -Raw | ConvertFrom-Json
$app = $null
try {
  $word = $null
  try { $word = New-Object -ComObject Word.Application } catch {}
  if (-not $word) {
    try { $word = New-Object -ComObject KWPS.Application } catch {}
  }
  if (-not $word) { Write-Output 'ERR=NO_OFFICE'; exit 3 }
  $app = $word
  try { $app.Visible = $false } catch {}
  try { $app.DisplayAlerts = 0 } catch {}
  $doc = $app.Documents.Open($job.input, $false, $true)
  $pages = $doc.ComputeStatistics(2)
  Write-Output ('PAGES=' + $pages)
  $ok = $false
  try {
    $doc.ExportAsFixedFormat($job.output, 17, $false, 0, 0, 1, $pages, 0, $true, $true, 1, $true, $false, $false)
    $ok = $true
  } catch { }
  if (-not $ok) {
    try { $doc.ExportPdf($job.output); $ok = $true } catch {}
  }
  if (-not $ok) { try { $doc.SaveAs([ref]$job.output, [ref]17); $ok = $true } catch {} }
  Write-Output ('EXPORTED=' + $ok)
  try { $doc.Close(0) } catch {}
} catch {
  Write-Output ('ERR=' + $_.Exception.Message)
} finally {
  if ($app) { try { $app.Quit() } catch {} }
}
"""


def have_pypdf():
    try:
        import pypdf  # noqa: F401
        return True
    except Exception:
        return False


def export_with_office(docx_path, pdf_path):
    """调 Word/WPS 把整本 docx 导成 pdf，返回 (成功, 页数, 错误)。"""
    pw = shutil.which('powershell') or shutil.which('pwsh')
    if not pw:
        return False, 0, '找不到 PowerShell，无法调用 Word/WPS'

    tmpdir = tempfile.mkdtemp(prefix='rz_pdf_')
    job = {'input': os.path.abspath(docx_path), 'output': os.path.abspath(pdf_path)}
    job_path = os.path.join(tmpdir, 'job.json')
    with open(job_path, 'w', encoding='utf-8') as f:
        json.dump(job, f, ensure_ascii=False)

    script = PS_TEMPLATE.replace('__JOB__', job_path.replace('\\', '\\\\'))
    ps_path = os.path.join(tmpdir, 'run.ps1')
    # 纯 ASCII 脚本 + UTF-8 BOM，路径通过 job.json 传入，避免中文路径在编码上出问题
    with open(ps_path, 'w', encoding='utf-8-sig') as f:
        f.write(script)

    pages, err = 0, ''
    try:
        r = subprocess.run([pw, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                            '-File', ps_path],
                           capture_output=True, text=True, timeout=900)
        for line in (r.stdout or '').splitlines():
            line = line.strip()
            if line.startswith('PAGES='):
                try:
                    pages = int(line[6:])
                except ValueError:
                    pass
            elif line.startswith('ERR='):
                err = line[4:]
            elif line.startswith('EXPORTED=False'):
                err = err or '导出未成功'
        if not err and not os.path.isfile(pdf_path):
            err = (r.stderr or '导出的 PDF 未生成').strip()[-400:]
    except subprocess.TimeoutExpired:
        err = '导出超时（900 秒）'
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    ok = os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0
    if not ok and 'NO_OFFICE' in err:
        ok, pages, err = _soffice_export(docx_path, pdf_path)
    return ok, pages, ('' if ok else (err or '未知错误'))


def _soffice_export(docx_path, pdf_path):
    """没装 Word/WPS 时用 LibreOffice 兜底（需本机已装，不下载）。"""
    exe = None
    for cand in ('C:/Program Files/LibreOffice/program/soffice.exe',
                 'C:/Program Files (x86)/LibreOffice/program/soffice.exe'):
        if os.path.isfile(cand):
            exe = cand
            break
    if not exe:
        return False, 0, '本机没有 Word / WPS / LibreOffice，无法自动出 PDF，请在 Word 里手动另存'
    try:
        r = subprocess.run([exe, '--headless', '--convert-to', 'pdf',
                            '--outdir', os.path.dirname(os.path.abspath(pdf_path)),
                            os.path.abspath(docx_path)],
                           capture_output=True, text=True, timeout=600)
        produced = os.path.join(os.path.dirname(os.path.abspath(pdf_path)),
                                os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
        if os.path.isfile(produced):
            if os.path.abspath(produced) != os.path.abspath(pdf_path):
                os.replace(produced, pdf_path)
            return True, count_pages(pdf_path), ''
        return False, 0, (r.stderr or 'LibreOffice 转换失败').strip()[-300:]
    except Exception as e:
        return False, 0, 'LibreOffice 调用失败：%s' % e


def slice_pdf(src, dst, first, last):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(src)
    total = len(reader.pages)
    first = max(1, first)
    last = total if last in (0, None) else min(last, total)
    writer = PdfWriter()
    for i in range(first - 1, last):
        writer.add_page(reader.pages[i])
    with open(dst, 'wb') as f:
        writer.write(f)
    return total, last - first + 1


def merge_pdfs(output, inputs):
    from pypdf import PdfReader, PdfWriter
    writer = PdfWriter()
    for path in inputs:
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)
    with open(output, 'wb') as f:
        writer.write(f)
    return sum(len(PdfReader(p).pages) for p in inputs)


def count_pages(pdf_path):
    try:
        from pypdf import PdfReader
        return len(PdfReader(pdf_path).pages)
    except Exception:
        return 0


def main():
    ap = argparse.ArgumentParser(description='docx 转 PDF（本机 Word/WPS/LibreOffice）与 PDF 截页合并')
    ap.add_argument('--input', help='源 docx')
    ap.add_argument('--output', help='目标 pdf')
    ap.add_argument('--pages', default=None, help='页范围，如 1-30')
    ap.add_argument('--last', type=int, default=None, help='取最后 N 页')
    ap.add_argument('--merge', nargs='+', default=None,
                    help='合并：--merge 输出.pdf 输入1.pdf 输入2.pdf')
    ap.add_argument('--info', action='store_true', help='只打印 docx 页数')
    args = ap.parse_args()

    if args.merge:
        if len(args.merge) < 3:
            sys.exit('用法：--merge 输出.pdf 输入1.pdf 输入2.pdf ...')
        out, ins = args.merge[0], args.merge[1:]
        if not have_pypdf():
            sys.exit('合并 PDF 需要 pypdf：python bootstrap.py --install')
        n = merge_pdfs(out, ins)
        print('已合并 %d 个文件 → %s（共 %d 页）' % (len(ins), out, n))
        return

    if not args.input:
        sys.exit('缺少 --input')
    if not os.path.isfile(args.input):
        sys.exit('文件不存在：%s' % args.input)

    want_pages = args.pages is not None or args.last is not None
    out = args.output or os.path.splitext(args.input)[0] + '.pdf'

    if args.info:
        tmpdir = tempfile.mkdtemp(prefix='rz_info_')
        tmp = os.path.join(tmpdir, 'probe.pdf')
        ok, pages, err = export_with_office(args.input, tmp)
        shutil.rmtree(tmpdir, ignore_errors=True)
        if not ok:
            sys.exit('无法读取页数：%s' % err)
        print('页数：%d' % pages)
        return

    if not want_pages:
        ok, pages, err = export_with_office(args.input, out)
        if not ok:
            sys.exit('导出失败：%s' % err)
        print('已导出：%s（%s 页）' % (out, pages or count_pages(out)))
        return

    # 需要截页：先整本导出，再用 pypdf 精确切，避免依赖 Word 的 Range 参数
    tmpdir = tempfile.mkdtemp(prefix='rz_slice_')
    full = os.path.join(tmpdir, 'full.pdf')
    ok, pages, err = export_with_office(args.input, full)
    if not ok:
        shutil.rmtree(tmpdir, ignore_errors=True)
        sys.exit('导出失败：%s' % err)
    if not pages:
        pages = count_pages(full)

    if args.pages:
        try:
            a, b = args.pages.replace('－', '-').split('-')
            first, last = int(a), int(b)
        except ValueError:
            shutil.rmtree(tmpdir, ignore_errors=True)
            sys.exit('--pages 格式应为 1-30')
    else:
        first, last = pages - args.last + 1, pages

    if not have_pypdf():
        shutil.rmtree(tmpdir, ignore_errors=True)
        sys.exit('截页需要 pypdf：python bootstrap.py --install（或直接在 Word 里导出指定页范围）')

    total, got = slice_pdf(full, out, first, last)
    shutil.rmtree(tmpdir, ignore_errors=True)
    print('已截取第 %d–%d 页 → %s（原文档共 %d 页，实际 %d 页）' % (first, last, out, total, got))


if __name__ == '__main__':
    main()
