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
RADIUS = 185 / 1113.0        # corner radius as a fraction of tile width


def tile_bbox(img):
    """Bounding box of the blue tile, ignoring the white ground and its shadow.

    The threshold has to stay generous on brightness: the tile's own shaded
    bottom edge falls to b~80, and excluding it clipped ~10px off the bottom,
    which pulled the crop upward and tilted the rounded mask off the artwork's
    real corners.  The drop shadow is neutral grey, so the b-r term alone is
    enough to reject it.
    """
    a = np.asarray(img.convert('RGB')).astype(int)
    r, b = a[..., 0], a[..., 2]
    ys, xs = np.nonzero((b - r > 10) & (b > 50))
    return xs.min(), ys.min(), xs.max(), ys.max()


def square_crop(img):
    """Crop tight to the tile.

    The source tile is 1113x1108 -- not quite square.  Padding it out to a
    square left uneven slivers of white above and below, so the corner mask no
    longer matched the artwork's corners.  Cropping tight and letting render()
    scale to a square instead spreads that 0.45% over the whole height, which
    is invisible, and keeps the tile flush to all four edges.
    """
    x0, y0, x1, y1 = tile_bbox(img)
    return img.crop((x0, y0, x1 + 1, y1 + 1))


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
