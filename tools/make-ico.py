"""Second half of the icon build (see tools/make-icons.js): copies the rendered PNGs from tests/out into icons/
and packs the Windows sizes into desktop/icon.ico (used by the EXE, its windows, the installer and PDF files in Explorer).
Needs Pillow:  python tools/make-ico.py"""
import os
from PIL import Image

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out = os.path.join(root, 'tests', 'out')
for name in ['icon-192.png', 'icon-512.png', 'icon-512-maskable.png', 'apple-touch-icon.png']:  # (lossless; a 256-colour palette bands the gradient)
    Image.open(os.path.join(out, name)).convert('RGBA').save(os.path.join(root, 'icons', name), 'PNG', optimize=True)
sizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
frames = [Image.open(os.path.join(out, f'ico-{n}.png')).convert('RGBA') for n in sizes]
frames[-1].save(os.path.join(root, 'desktop', 'icon.ico'), format='ICO', sizes=[(n, n) for n in sizes], append_images=frames[:-1])
print('icons/ updated, desktop/icon.ico written with sizes', sizes)
