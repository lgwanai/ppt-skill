#!/usr/bin/env python3
"""将 PPT/PPTX 每一页转为 PNG 图片，使用 LibreOffice + Python"""

import subprocess, sys, os, shutil, tempfile, glob

def ppt_to_png(ppt_path, output_dir=None):
    ppt_path = os.path.abspath(ppt_path)
    if not os.path.exists(ppt_path):
        print(f'文件不存在: {ppt_path}')
        return
    
    base = os.path.splitext(os.path.basename(ppt_path))[0]
    if output_dir is None:
        output_dir = f'{base}_png'
    os.makedirs(output_dir, exist_ok=True)
    
    print(f'🔧 LibreOffice 转换 PDF...')
    tmpdir = tempfile.mkdtemp()
    try:
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', tmpdir, ppt_path
        ], check=True, timeout=120)
    except subprocess.TimeoutExpired:
        print('超时, 重试...')
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            '--outdir', tmpdir, ppt_path
        ], check=True, timeout=300)
    except FileNotFoundError:
        # Try libreoffice variant
        for cmd in ['soffice', 'libreoffice']:
            try:
                subprocess.run([
                    cmd, '--headless', '--convert-to', 'pdf',
                    '--outdir', tmpdir, ppt_path
                ], check=True, timeout=120)
                break
            except FileNotFoundError:
                continue
        else:
            print('未找到 LibreOffice. 请安装: brew install libreoffice')
            return
    
    pdf_files = glob.glob(os.path.join(tmpdir, '*.pdf'))
    if not pdf_files:
        print('PDF 生成失败')
        return
    pdf_path = pdf_files[0]
    print(f'   PDF 已生成: {pdf_path}')
    
    print(f'🖼️  PDF → PNG (每页)...')
    try:
        # Try PyMuPDF (fastest)
        import fitz
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=200)
            out = os.path.join(output_dir, f'{base}_p{i+1:02d}.png')
            pix.save(out)
            print(f'   ✅ {out}')
        doc.close()
    except ImportError:
        try:
            # Try pdf2image (poppler-based)
            from pdf2image import convert_from_path
            images = convert_from_path(pdf_path, dpi=200)
            for i, img in enumerate(images):
                out = os.path.join(output_dir, f'{base}_p{i+1:02d}.png')
                img.save(out)
                print(f'   ✅ {out}')
        except ImportError:
            # Try PIL-only (slower, less accurate)
            try:
                from PIL import Image
                subprocess.run([
                    'python3', '-c', 
                    f'''from PIL import Image; import fitz; doc=fitz.open("{pdf_path}");
[doc[i].get_pixmap(dpi=150).save("{os.path.join(output_dir, base)}_p%02d.png" % (i+1)) for i in range(len(doc))]'''
                ])
            except:
                print('请安装: pip install pymupdf    (推荐)')
                print('或安装: pip install pdf2image && brew install poppler')
                return
    
    shutil.rmtree(tmpdir, ignore_errors=True)
    print(f'\n完成! {output_dir}/ 目录下有 {len(os.listdir(output_dir))} 张图片')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python ppt_to_png.py <文件.pptx> [输出目录]')
        sys.exit(1)
    ppt_to_png(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
