#!/usr/bin/env python3
"""PPTX Notes Exporter: tablet-friendly PDFs, Markdown notes, and slide PNG ZIP.

Install: py -m pip install python-pptx reportlab pywin32
Run: py pptx_notes_exporter.py presentation.pptx
PowerPoint on Windows is used for thumbnails; --no-thumbnails skips it.
"""
import argparse

__version__ = "1.0.1"
import re
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from xml.sax.saxutils import escape

from pptx import Presentation
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

NAVY = colors.HexColor('#153552')
INK = colors.HexColor('#152434')
MUTED = colors.HexColor('#64758a')


def fonts():
    # Malgun Gothic covers Latin, Hangul, and CJK ideographs (緣 / 縁 / 缘).
    choices = [
        ('C:/Windows/Fonts/malgun.ttf', 'C:/Windows/Fonts/malgunbd.ttf'),
        ('/usr/share/fonts/truetype/nanum/NanumGothic.ttf', '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'),
        ('C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/msyhbd.ttc'),
        ('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
         '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'),
    ]
    # Register Verdana separately for Latin text; keep CJK fallback choices below.
    latin_regular = Path('C:/Windows/Fonts/verdana.ttf')
    latin_bold = Path('C:/Windows/Fonts/verdanab.ttf')
    if latin_regular.exists():
        pdfmetrics.registerFont(TTFont('LatinReader', str(latin_regular)))
        pdfmetrics.registerFont(TTFont('LatinReaderBold', str(latin_bold if latin_bold.exists() else latin_regular)))
        pdfmetrics.registerFontFamily('LatinReader', normal='LatinReader', bold='LatinReaderBold')
        print('Latin font: Verdana Regular')
    else:
        print('Verdana not found: using CJK fallback font for all text.')
    for regular, bold in choices:
        if Path(regular).exists():
            try:
                pdfmetrics.registerFont(TTFont('Reader', regular, subfontIndex=0))
                if Path(bold).exists():
                    pdfmetrics.registerFont(TTFont('ReaderBold', bold, subfontIndex=0))
                else:
                    pdfmetrics.registerFont(TTFont('ReaderBold', regular, subfontIndex=0))
                pdfmetrics.registerFontFamily('Reader', normal='Reader', bold='ReaderBold')
                return 'Reader', 'ReaderBold'
            except Exception as exc:
                print(f'Font {regular} unavailable: {exc}')
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    pdfmetrics.registerFontFamily('STSong-Light', normal='STSong-Light', bold='STSong-Light')
    return 'STSong-Light', 'STSong-Light'


def extract(pptx):
    prs = Presentation(str(pptx))
    result = []
    for n, slide in enumerate(prs.slides, 1):
        title = slide.shapes.title.text.strip() if slide.shapes.title and slide.shapes.title.has_text_frame else ''
        if not title:
            # Prefer topmost, substantial text box when there is no explicit title placeholder.
            candidates = [(s.top, s.text.strip()) for s in slide.shapes
                          if s.has_text_frame and s.text.strip() and len(s.text.strip()) < 180]
            title = sorted(candidates)[0][1] if candidates else 'Untitled slide'
        title = ' '.join(title.split())
        if not slide.has_notes_slide or not slide.notes_slide.notes_text_frame:
            continue
        frame = slide.notes_slide.notes_text_frame
        lines = []
        for paragraph in frame.paragraphs:
            if not paragraph.text.strip() or paragraph.text.strip() == str(n):
                continue
            # Preserve character-level bold from PowerPoint notes runs.
            # A run with no explicit bold setting follows the paragraph default.
            pieces = []
            for run in paragraph.runs:
                if not run.text:
                    continue
                is_bold = run.font.bold
                if is_bold is None:
                    is_bold = paragraph.font.bold
                pieces.append(('**' + run.text + '**') if is_bold else run.text)
            lines.append(''.join(pieces).strip())
        if lines:
            result.append((n, title, lines))
    return result


def thumbnails(pptx, tmp):
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print('No pywin32: thumbnails omitted.')
        return {}
    app = pres = None
    imgs = {}
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx('PowerPoint.Application')
        pres = app.Presentations.Open(str(pptx.resolve()), ReadOnly=True,
                                      Untitled=False, WithWindow=False)
        for i in range(1, pres.Slides.Count + 1):
            path = Path(tmp) / f'slide_{i:03d}.png'
            pres.Slides(i).Export(str(path.resolve()), 'PNG', 960, 540)
            if path.exists():
                imgs[i] = path
    except Exception as exc:
        print(f'Thumbnails unavailable: {exc}')
    finally:
        if pres is not None:
            try: pres.Close()
            except Exception: pass
        if app is not None:
            try: app.Quit()
            except Exception: pass
        pythoncom.CoUninitialize()
    return imgs


def mark(text):
    """Verdana for Latin runs; CJK glyphs stay in Reader fallback.

    Existing **bold** and transition-word emphasis are preserved.
    """
    latin_available = 'LatinReader' in pdfmetrics.getRegisteredFontNames()
    parts = re.split(r'(\*\*.*?\*\*)', text)
    out = []
    for part in parts:
        emphasized = part.startswith('**') and part.endswith('**')
        if emphasized:
            part = part[2:-2]
        chunks = re.split(r'([\u0080-\uffff]+)', part)
        rendered = []
        for chunk in chunks:
            if not chunk:
                continue
            safe = escape(chunk)
            if latin_available and not re.search(r'[\u0080-\uffff]', chunk):
                safe = '<font name="' + ('LatinReaderBold' if emphasized else 'LatinReader') + '">' + safe + '</font>'
            if not emphasized:
                safe = re.sub(r'\b(First|Second|Third|Finally)\b', r'<b>\1</b>', safe)
            rendered.append(safe)
        joined = ''.join(rendered)
        out.append(joined if not emphasized else '<b>' + joined + '</b>')
    return ''.join(out)


def para(text, font, size, leading=None, bold=False):
    style = ParagraphStyle('x', fontName=('ReaderBold' if font == 'Reader' else font) if bold else font,
                           fontSize=size, leading=leading or size*1.38,
                           textColor=INK, splitLongWords=True)
    return Paragraph(mark(text), style)


def draw_pdf(output, notes, imgs, orient, regular, bold, min_font=11.5, aspect="4:3", base_font=22.5, max_font=24.0):
    """One page per slide; text flows beside then below the slide thumbnail."""
    # Keep the original long side (800 pt) for screen presets.
    # A4 uses its actual physical dimensions in points.
    if aspect == "A4":
        W, H = (841.89, 595.28) if orient == "landscape" else (595.28, 841.89)
    else:
        a, b = map(int, aspect.split(":"))
        short = 800 * b / a
        W, H = (800, short) if orient == "landscape" else (short, 800)
    c = canvas.Canvas(str(output), pagesize=(W, H))
    c.setTitle('KOBIC - presenter notes - ' + orient)
    margin = 25 if orient == 'landscape' else 28
    bar_h = 48 if orient == 'landscape' else 52
    footer_y = 15
    bottom = 32
    thumb_w = 150 if orient == 'landscape' else 142
    thumb_h = thumb_w * 9 / 16
    thumb_x = W - margin - thumb_w
    thumb_top = H - bar_h - 12
    thumb_bottom = thumb_top - thumb_h
    gutter = 15
    full_w = W - 2 * margin
    narrow_w = thumb_x - margin - gutter
    top = H - bar_h - 14

    for number, title, lines in notes:
        # Flow text one *rendered line* at a time.  A line whose bottom would
        # overlap the thumbnail uses the narrow width; the very next line can
        # use the full page width, even in the middle of a paragraph.
        # Paragraph.split() preserves inline font tags and bold run formatting.
        def compose(size, draw=False):
            y = top
            leading = size * 1.43
            space = max(5, size * .40)
            style = ParagraphStyle(
                'notes', fontName=regular, fontSize=size,
                leading=leading, textColor=INK, splitLongWords=True,
                allowWidows=1, allowOrphans=1)
            for line in lines:
                if not line.strip():
                    continue
                remaining = Paragraph(mark(line.strip()), style)
                while remaining is not None:
                    if y - leading < bottom - .01:
                        return False, y
                    # Choose the width for this individual line, not for the
                    # whole paragraph. This eliminates the large blank strip
                    # created by jumping to thumb_bottom after a split.
                    # More simply: any line that intersects the image's
                    # vertical extent must be narrow.
                    if y > thumb_bottom + 0.01:
                        width = narrow_w
                    else:
                        width = full_w
                    remaining.wrap(width, 10000)
                    if remaining.height <= leading + .01:
                        first, rest = remaining, None
                    else:
                        # ReportLab cannot split off a single line reliably;
                        # permit single-line splits by disabling widow/orphan restrictions.
                        parts = remaining.split(width, leading + .01)
                        if len(parts) < 2:
                            return False, y
                        first, rest = parts[0], parts[1]
                    _, height = first.wrap(width, 10000)
                    if y - height < bottom - .01:
                        return False, y
                    if draw:
                        first.drawOn(c, margin, y - height)
                    y -= height
                    remaining = rest
                y -= space
            return y >= bottom - .01, y

        # Typography is intentionally stable across slides.  Use the base size
        # unless a slide is genuinely short; only then move up by 1 or 1.5 pt.
        # Longer notes may shrink as needed, but never enlarge every slide to
        # fill the page.  The remaining white space is useful while presenting.
        base = min(base_font, max_font)
        base_fits, base_end = compose(base)
        if base_fits:
            free_fraction = max(0.0, base_end - bottom) / (top - bottom)
            if free_fraction >= 0.38:
                desired = min(max_font, base + 1.5)
            elif free_fraction >= 0.24:
                desired = min(max_font, base + 1.0)
            else:
                desired = base
        else:
            desired = base

        size = desired
        while size >= min_font - 0.001:
            fits, _ = compose(size)
            if fits:
                break
            size = round(size - 0.25, 3)
        if size < min_font - 0.001:
            raise RuntimeError(
                f'Slide {number}: notes cannot fit on one page even at '
                f'{min_font} pt. Shorten notes or reduce --min-font.')

        c.setFillColor(NAVY)
        c.rect(0, H - bar_h, W, bar_h, fill=1, stroke=0)
        # One-line heading: fixed-size No.## + dynamically fitted title.
        # Never wrap or truncate the title; shrink its font only if needed.
        c.setFillColor(colors.white)
        number_label = f'No.{number:02d}'
        number_size = 19 if orient == 'landscape' else 18
        baseline = H - bar_h / 2 - number_size * .34
        c.setFont(bold, number_size)
        c.drawString(margin, baseline, number_label)
        title_x = margin + pdfmetrics.stringWidth(number_label, bold, number_size) + 17
        title_available = W - margin - title_x
        title_size = 29 if orient == 'landscape' else 26
        while title_size > 7 and pdfmetrics.stringWidth(title, bold, title_size) > title_available:
            title_size -= .25
        if pdfmetrics.stringWidth(title, bold, title_size) > title_available:
            raise RuntimeError(f'Slide {number}: title too long for a single line.')
        c.setFont(bold, title_size)
        c.drawString(title_x, baseline + (number_size - title_size) * .15, title)

        if number in imgs:
            c.drawImage(str(imgs[number]), thumb_x, thumb_bottom,
                        width=thumb_w, height=thumb_h,
                        preserveAspectRatio=True, anchor='c', mask='auto')
        else:
            c.setStrokeColor(colors.HexColor('#cbd5e1'))
            c.rect(thumb_x, thumb_bottom, thumb_w, thumb_h)
            c.setFillColor(MUTED)
            c.setFont(regular, 9)
            c.drawCentredString(thumb_x + thumb_w / 2,
                                thumb_bottom + thumb_h / 2, 'SLIDE PREVIEW')
        fits, _ = compose(size, draw=True)
        assert fits
        c.setStrokeColor(colors.HexColor('#dce4ed'))
        c.line(margin, 26, W - margin, 26)
        c.setFillColor(MUTED)
        c.setFont(regular, 8)
        c.drawString(margin, footer_y, 'CJK Bioinformatics Symposium  ·  November 17, 2026')
        c.drawRightString(W - margin, footer_y, f'{number} / {len(notes)}')
        c.showPage()
    c.save()
    return len(notes)


def export_notes_markdown(path, notes, source_name):
    """Export all speaker notes as a readable, UTF-8 Markdown document."""
    def md_heading(text):
        return text.replace('\n', ' ').strip()

    chunks = [f'# Presenter Notes — {source_name}', '',
              f'*{len(notes)} slides with speaker notes.*', '']
    for number, title, lines in notes:
        chunks.extend([f'## Slide {number:02d} — {md_heading(title)}', ''])
        # Each PowerPoint note paragraph becomes its own Markdown paragraph.
        # Bold markers created during extraction are retained.
        for line in lines:
            chunks.extend([line.strip(), ''])
    path.write_text('\n'.join(chunks).rstrip() + '\n', encoding='utf-8-sig')
    return path


def archive_slide_images(path, imgs):
    """Store the actual exported slide PNGs, without resizing or recompression."""
    if not imgs:
        return 0
    with ZipFile(path, 'w', compression=ZIP_DEFLATED, compresslevel=6) as archive:
        for number, png_path in sorted(imgs.items()):
            archive.write(png_path, arcname=f'slide_{number:03d}.png')
    return len(imgs)


def main():
    ap = argparse.ArgumentParser(
        description='PPTX Notes Exporter: export speaker notes to tablet-friendly PDFs and Markdown, plus a slide PNG ZIP.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        prog="pptx_notes_exporter.py",
        epilog=f"""Version: {__version__}

Examples:
  python pptx_notes_exporter.py presentation.pptx
      Default: 4:3 landscape + portrait PDFs for classic iPads.

  python pptx_notes_exporter.py presentation.pptx --aspect 16:10
      Widescreen laptop/tablet reading.

  python pptx_notes_exporter.py presentation.pptx --aspect A4
      Landscape and portrait A4 printing.

  python pptx_notes_exporter.py presentation.pptx --no-thumbnails
      Skip PowerPoint slide previews and the PNG ZIP (no PowerPoint needed).
      The Markdown notes file is still created.

  python pptx_notes_exporter.py presentation.pptx --output-dir output
      Save all generated outputs in the output folder.

  python pptx_notes_exporter.py presentation.pptx --base-font 23 --max-font 24
      Slightly larger body text while keeping slides visually consistent.

  python pptx_notes_exporter.py presentation.pptx --min-font 10
      Allow extra-dense notes to shrink further (use only if necessary).

  python pptx_notes_exporter.py presentation.pptx --no-image-zip
      Embed previews in the PDFs but do not retain the PNG archive.

  python pptx_notes_exporter.py presentation.pptx --no-markdown
      Create the PDFs and image ZIP without the Markdown transcript.

Page formats:
  4:3      Classic iPad, e.g. iPad 3 (default)
  16:10    Widescreen laptops and tablets (not A4)
  A4       ISO A4 paper, landscape and portrait

Typography:
  Base body font: 22.5 pt; short slides may grow to 23.5-24 pt.
  Dense slides shrink only as needed to fit one slide per page.
  Line spacing scales with font size (1.43x); thumbnails stay small.
  Original notes bold runs are retained where PPTX formatting permits.
  Titles start at 29 pt (landscape) or 26 pt (portrait), shrinking to one line.

Output (both orientations, files in PPTX folder by default):
  <presentation>_notes_8inch_landscape.pdf
  <presentation>_notes_8inch_portrait.pdf
  <presentation>_notes.md                 Complete speaker notes, UTF-8
  <presentation>_slides_png.zip           Exported 960x540 slide PNGs
  Re-running with another --aspect overwrites PDFs with the same names.
  The ZIP is produced only when PowerPoint slide export succeeds.
  Slides with no notes are still included in the image ZIP.

Requirements:
  pip install python-pptx reportlab
  pip install pywin32     # optional, for thumbnails on Windows with PowerPoint
""")
    ap.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    ap.add_argument('pptx', type=Path, help='Input PowerPoint .pptx file')
    ap.add_argument('--output-dir', type=Path, help='Destination folder (default: PPTX folder)')
    ap.add_argument('--no-thumbnails', action='store_true',
                    help='Do not export slide previews')
    ap.add_argument('--no-image-zip', action='store_true',
                    help='Do not save exported slide PNGs as a ZIP archive')
    ap.add_argument('--no-markdown', action='store_true',
                    help='Do not export speaker notes to a Markdown (.md) file')
    ap.add_argument('--aspect', choices=('4:3', '16:10', 'A4'), default='4:3',
                    help='Page ratio / paper format (default: 4:3)')
    ap.add_argument('--base-font', type=float, default=22.5,
                    help='Normal body font size in points (default: 22.5)')
    ap.add_argument('--max-font', type=float, default=24.0,
                    help='Maximum body font size for short notes (default: 24)')
    ap.add_argument('--min-font', type=float, default=11.5,
                    help='Minimum body font size for dense notes (default: 11.5)')
    args = ap.parse_args()
    if not (0 < args.min_font <= args.base_font <= args.max_font):
        ap.error('Font sizes must satisfy 0 < --min-font <= --base-font <= --max-font')
    if not args.pptx.is_file():
        ap.error('PPTX file not found: '+str(args.pptx))
    regular,bold = fonts()
    notes = extract(args.pptx)
    if not notes:
        ap.error('No speaker notes found.')
    dest = args.output_dir or args.pptx.parent
    dest.mkdir(parents=True,exist_ok=True)
    if not args.no_markdown:
        md_path = dest / f'{args.pptx.stem}_notes.md'
        export_notes_markdown(md_path, notes, args.pptx.name)
        print(f'Created {md_path} (Markdown notes)')
    with tempfile.TemporaryDirectory() as tmp:
        imgs = {} if args.no_thumbnails else thumbnails(args.pptx,tmp)
        if not args.no_image_zip:
            if imgs:
                zip_path = dest / f'{args.pptx.stem}_slides_png.zip'
                archived = archive_slide_images(zip_path, imgs)
                print(f'Created {zip_path} ({archived} slide PNGs)')
            else:
                print('Slide PNG ZIP skipped: no images were exported.')
        for orientation in ('landscape','portrait'):
            path = dest / f'{args.pptx.stem}_notes_8inch_{orientation}.pdf'
            count = draw_pdf(path,notes,imgs,orientation,regular,bold,
                             min_font=args.min_font, aspect=args.aspect,
                             base_font=args.base_font, max_font=args.max_font)
            print(f'Created {path} ({count} pages)')
    print(f'{len(notes)} slides with notes.')

if __name__=='__main__':
    main()
