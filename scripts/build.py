"""Package Word Helper for the Chrome Web Store.

    python scripts/build.py

Writes dist/word-helper-<version>.zip containing only what the extension runs:
the manifest and every file it references, plus the pages and assets those load.
README, LICENSE, source artwork and tooling are left out.  The file list is
derived from manifest.json, so a newly referenced file can't be forgotten; the
script fails if anything referenced is missing, and refuses to package a build
with debug logging switched on.
"""
import json
import os
import re
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def manifest_files(m):
    files = {'manifest.json'}
    files.update(m.get('icons', {}).values())
    action = m.get('action', {})
    files.update(action.get('default_icon', {}).values())
    if 'default_popup' in action:
        files.add(action['default_popup'])
    if 'service_worker' in m.get('background', {}):
        files.add(m['background']['service_worker'])
    for cs in m.get('content_scripts', []):
        files.update(cs.get('js', []))
        files.update(cs.get('css', []))
    for war in m.get('web_accessible_resources', []):
        files.update(war.get('resources', []))
    return files


def page_files(page):
    """Scripts, stylesheets and images an extension page loads by relative path."""
    with open(os.path.join(ROOT, page), encoding='utf-8') as f:
        html = f.read()
    refs = re.findall(r'''(?:src|href)=["']([^"':]+)["']''', html)
    return {r for r in refs if not r.startswith(('#', '//'))}


def main():
    with open(os.path.join(ROOT, 'manifest.json'), encoding='utf-8') as f:
        manifest = json.load(f)

    files = manifest_files(manifest)
    for page in [p for p in files if p.endswith('.html')]:
        files |= page_files(page)

    missing = sorted(p for p in files if not os.path.isfile(os.path.join(ROOT, p)))
    if missing:
        sys.exit('missing files referenced by the manifest: ' + ', '.join(missing))

    with open(os.path.join(ROOT, 'content.js'), encoding='utf-8') as f:
        if re.search(r'^const DEBUG = true;', f.read(), re.M):
            sys.exit('content.js has DEBUG = true; switch it off before packaging')

    out_dir = os.path.join(ROOT, 'dist')
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"word-helper-{manifest['version']}.zip")
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in sorted(files):
            z.write(os.path.join(ROOT, p), p)

    print(f'wrote {os.path.relpath(out, ROOT)}  ({os.path.getsize(out) / 1024:.1f} KB)')
    for p in sorted(files):
        print(f'  {p:18} {os.path.getsize(os.path.join(ROOT, p)) / 1024:7.1f} KB')


if __name__ == '__main__':
    main()
