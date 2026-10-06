"""Builds the small sample PDFs in tests/samples used by tests/selftest.js.  Needs: pip install reportlab pillow
   Run from the project folder:  python tests/make_samples.py"""
import io, os, random
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__)); OUT = os.path.join(HERE, 'samples'); os.makedirs(OUT, exist_ok=True)
LIB = os.path.join(HERE, '..', 'lib', 'fonts'); W, H = 595, 842
def font(name, size): return ImageFont.truetype(os.path.join(LIB, name), size)
pdfmetrics.registerFont(TTFont('Lato', os.path.join(LIB, 'Lato-Regular.ttf'))); pdfmetrics.registerFont(TTFont('Lato-Bold', os.path.join(LIB, 'Lato-Bold.ttf'))); pdfmetrics.registerFont(TTFont('Lora', os.path.join(LIB, 'Lora-Regular.ttf')))

# 1. native text: embedded TrueType (regular + a bold word), standard fonts, white on a colour band, rotated text, a picture with words in it
c = canvas.Canvas(os.path.join(OUT, 'native.pdf'), pagesize=(W, H))
c.setFont('Lato', 14); c.drawString(60, 780, 'Invoice number 2048 for '); x = 60 + pdfmetrics.stringWidth('Invoice number 2048 for ', 'Lato', 14)
c.setFont('Lato-Bold', 14); c.drawString(x, 780, 'Acme Trading'); x += pdfmetrics.stringWidth('Acme Trading', 'Lato-Bold', 14)
c.setFont('Lato', 14); c.drawString(x, 780, ' dated today')
c.setFont('Lora', 12)
for i, l in enumerate(['The quick brown fox jumps over the lazy dog and keeps', 'running through the quiet forest until the evening light', 'fades behind the distant hills.']): c.drawString(60, 730 - i * 16, l)
c.setFont('Times-Roman', 13); c.drawString(60, 650, 'Standard Times text that is not embedded')
c.setFont('Helvetica-Bold', 13); c.drawString(60, 625, 'Standard Helvetica bold heading')
c.setFillColorRGB(0.1, 0.3, 0.6); c.rect(50, 560, 400, 40, fill=1, stroke=0); c.setFillColorRGB(1, 1, 1); c.setFont('Lato-Bold', 18); c.drawString(60, 573, 'White text on a blue band')
c.setFillColorRGB(0, 0, 0); c.saveState(); c.translate(520, 300); c.rotate(90); c.setFont('Lato', 12); c.drawString(0, 0, 'Rotated side label'); c.restoreState()
im = Image.new('RGB', (900, 300), (245, 238, 220)); d = ImageDraw.Draw(im)
d.text((40, 40), 'Text inside a picture', font=font('Lato-Bold.ttf', 64), fill=(120, 20, 20)); d.text((40, 160), 'Second line in the image 123', font=font('Lora-Regular.ttf', 52), fill=(20, 20, 20))
c.drawImage(ImageReader(im), 60, 380, width=300, height=100)
t = c.beginText(60, 330); t.setFont('Times-Roman', 12); t.setWordSpace(6); t.textOut('This line is justified with wide word spacing'); c.drawText(t)
c.bookmarkPage('top'); c.addOutlineEntry('Invoice', 'top', level=0)
c.showPage(); c.setFont('Helvetica', 12); c.drawString(60, 780, 'Second page'); c.bookmarkPage('p2'); c.addOutlineEntry('Second page', 'p2', level=0); c.showPage(); c.save()

# 2. a scan (picture of a page with light noise), 3. the same with a crooked tilt, 4. the same with a hidden OCR text layer
def scan_image(tilt=0):
    im = Image.new('RGB', (1654, 2339), (250, 248, 240)); d = ImageDraw.Draw(im)
    d.text((150, 160), 'SCANNED AGREEMENT', font=font('Lato-Bold.ttf', 70), fill=(25, 25, 25)); y = 330
    for l in ['This agreement is made on the first day of March', 'between the supplier and the customer named below.', 'All payments are due within thirty days of invoice.']:
        d.text((150, y), l, font=font('Lora-Regular.ttf', 46), fill=(30, 30, 30)); y += 64
    d.text((150, 640), 'Reference code: XD3-2026-0042', font=font('RobotoMono-Regular.ttf', 44), fill=(30, 30, 30))
    d.text((150, 760), 'Total amount payable: 12,500.00', font=font('Lato-Regular.ttf', 48), fill=(30, 30, 30))
    px = im.load(); random.seed(1)
    for _ in range(60000):
        x = random.randrange(1654); yy = random.randrange(2339); r, g, b = px[x, yy]; n = random.randint(-18, 6); px[x, yy] = (max(0, r + n), max(0, g + n), max(0, b + n))
    if tilt: im = im.rotate(tilt, resample=Image.BICUBIC, fillcolor=(250, 248, 240))
    return im.filter(ImageFilter.GaussianBlur(0.6))
for name, tilt in (('scan.pdf', 0), ('skew.pdf', 2.0)):
    b = io.BytesIO(); scan_image(tilt).save(b, 'JPEG', quality=80); b.seek(0)
    c = canvas.Canvas(os.path.join(OUT, name), pagesize=(W, H)); c.drawImage(ImageReader(b), 0, 0, width=W, height=H); c.showPage(); c.save()
b = io.BytesIO(); scan_image().save(b, 'JPEG', quality=80); b.seek(0); k = W / 1654
c = canvas.Canvas(os.path.join(OUT, 'searchable.pdf'), pagesize=(W, H)); c.drawImage(ImageReader(b), 0, 0, width=W, height=H)
for text, y, size in (('This agreement is made on the first day of March', 330 + 40, 46), ('Total amount payable: 12,500.00', 760 + 42, 48)):
    t = c.beginText(); t.setTextRenderMode(3); t.setFont('Helvetica', size * k * 1.05); t.setTextOrigin(150 * k, H - y * k); t.textOut(text); c.drawText(t)
c.showPage(); c.save()

# 5. a form: text fields, a checkbox and a dropdown
c = canvas.Canvas(os.path.join(OUT, 'form.pdf'), pagesize=(W, H)); f = c.acroForm
c.setFont('Helvetica', 12); c.drawString(60, 760, 'Name'); f.textfield(name='name', x=140, y=750, width=250, height=22, value='', fontSize=11, borderWidth=1)
c.drawString(60, 720, 'City'); f.textfield(name='city', x=140, y=710, width=250, height=22, value='Pune', fontSize=11, borderWidth=1)
c.drawString(60, 680, 'Agree'); f.checkbox(name='agree', x=140, y=672, size=18, checked=False, borderWidth=1)
c.drawString(60, 640, 'Plan'); f.choice(name='plan', x=140, y=630, width=160, height=22, options=['Basic', 'Standard', 'Premium'], value='Standard', borderWidth=1)
c.showPage(); c.save()
print('samples written to', OUT)
