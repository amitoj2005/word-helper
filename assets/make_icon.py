# Renders the Word Helper icon PNGs from the source artwork in assets/.
#
#   python assets/make_icon.py
#
# The source is a square-ish tile floating on a white background with a drop
# shadow.  We crop to the tile, replace the white surround with transparency
# using a rounded-rect mask cut to the tile's own corner radius, then
# downsample.  Without that mask the icon shows as a white square on dark
# toolbars and in Chrome's dark-mode extensions menu.
from PIL import Image, ImageDraw
import numpy as np, os

SRC    = 'icon-source.png'
SIZES  = (16, 48, 128)
RADIUS = 185 / 1109.0        # corner radius as a fraction of tile width


def tile_bbox(img):
    """Bounding box of the blue tile, ignoring the white ground and shadow."""
    a = np.asarray(img.convert('RGB')).astype(int)
    r, b = a[..., 0], a[..., 2]
    ys, xs = np.nonzero((b - r > 25) & (b > 90))
    return xs.min(), ys.min(), xs.max(), ys.max()


def square_crop(img):
    x0, y0, x1, y1 = tile_bbox(img)
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    half   = max(x1 - x0, y1 - y0) / 2
    return img.crop((round(cx - half), round(cy - half),
                     round(cx + half), round(cy + half)))


def render(size, tile):
    ss  = size * 8
    out = tile.convert('RGB').resize((size, size), Image.LANCZOS)
    # The mask is drawn at 8x and brought down with BOX (area average) rather
    # than LANCZOS, whose ringing can push edge alpha past 0/255.
    mask = Image.new('L', (ss, ss), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, ss - 1, ss - 1], radius=round(RADIUS * ss), fill=255)
    out = out.convert('RGBA')
    out.putalpha(mask.resize((size, size), Image.BOX))
    return out


if __name__ == '__main__':
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    tile = square_crop(Image.open(os.path.join(here, SRC)))
    print('tile cropped to', tile.size)
    for s in SIZES:
        render(s, tile).save(os.path.join(root, 'icon%d.png' % s))
        print('wrote icon%d.png' % s)
