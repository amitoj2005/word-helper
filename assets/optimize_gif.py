# Shrinks a raw screen recording into something a GitHub README will actually
# load.  GitHub caches READMEs through Camo, which refuses assets over ~10 MB,
# and a heavy GIF makes the page feel broken well before that.  Aim for ~2 MB.
#
#   python assets/optimize_gif.py docs/raw.gif docs/demo.gif
#   python assets/optimize_gif.py docs/raw.gif docs/demo.gif --width 800 --every 2
#
import argparse, os
from PIL import Image, ImageSequence


def optimize(src, dst, width, every, colors):
    im     = Image.open(src)
    frames = []
    for i, f in enumerate(ImageSequence.Iterator(im)):
        if i % every:                       # drop every Nth frame
            continue
        f = f.convert('RGBA')
        if f.width > width:
            h = round(f.height * width / f.width)
            f = f.resize((width, h), Image.LANCZOS)
        # ADAPTIVE keeps UI gradients from banding the way WEB would.
        frames.append(f.convert('RGB').quantize(colors=colors, method=Image.MEDIANCUT))

    if not frames:
        raise SystemExit('no frames read from %s' % src)

    # Frame delays are per-frame in the source; dropping frames means the kept
    # ones have to absorb the skipped time or playback runs fast.
    delay = im.info.get('duration', 50) * every

    frames[0].save(dst, save_all=True, append_images=frames[1:],
                   duration=delay, loop=0, optimize=True, disposal=2)

    mb = os.path.getsize(dst) / 1e6
    print('%s -> %s' % (src, dst))
    print('  %d frames (from %d), %dpx wide, %d colors, %.2f MB'
          % (len(frames), i + 1, frames[0].width, colors, mb))
    if mb > 3:
        print('  still heavy -- try --width 720 or --every 3')


if __name__ == '__main__':
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('src')
    p.add_argument('dst', nargs='?', default='docs/demo.gif')
    p.add_argument('--width',  type=int, default=900, help='max width (default 900)')
    p.add_argument('--every',  type=int, default=1,   help='keep 1 of every N frames')
    p.add_argument('--colors', type=int, default=128, help='palette size (default 128)')
    a = p.parse_args()
    optimize(a.src, a.dst, a.width, a.every, a.colors)
