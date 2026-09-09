"""Rendering pipeline shared by preview, pptx image insertion and docx image mode.

Primary path (highest fidelity, matches what Word/WPS render):
    LaTeX -> pandoc markdown $$..$$ -> temp .docx (OMML) -> LibreOffice headless
    (isolated -env:UserInstallation profile, never fights a user's live LO) -> PNG
    -> Pillow crop-tight + optional white->alpha transparency.

Degraded path (only when soffice is missing/broken, or pandoc itself is absent):
    LaTeX -> matplotlib mathtext PNG. mathtext covers only a subset of LaTeX
    (no cases/aligned environments); failures are reported as errors.
"""
import os
import subprocess


class LatexError(Exception):
    """LaTeX source rejected by pandoc — must NOT trigger the matplotlib fallback."""


def pandoc_available(tools):
    return bool(tools.get('pandoc')) and os.path.isfile(tools['pandoc'])


def soffice_available(tools):
    return bool(tools.get('soffice')) and os.path.isfile(tools['soffice'])


def _run(cmd, timeout):
    try:
        return subprocess.run(cmd, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise RuntimeError('command timed out after %ss: %s' % (timeout, cmd[0]))


def temp_docx_from_latex(latex, tmpdir, tools, timeout=60):
    """$$latex$$ markdown -> temp docx via pandoc. Raises LatexError on pandoc rejection."""
    md_path = os.path.join(tmpdir, 'f.md')
    docx_path = os.path.join(tmpdir, 'f.docx')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('$$\n' + latex + '\n$$\n')
    r = _run([tools['pandoc'], '-f', 'markdown', '-t', 'docx', '-o', docx_path, md_path], timeout)
    if r.returncode != 0 or not os.path.isfile(docx_path):
        # pandoc exits 43 on parse errors; surface a compact reason
        err = r.stderr.decode('utf-8', 'replace') if r.stderr else ''
        lines = [ln for ln in err.splitlines() if ln.strip()]
        reason = ' | '.join(lines[-2:]) if lines else ('exit code %s' % r.returncode)
        raise LatexError(reason[:400])
    return docx_path


def soffice_png_from_docx(docx_path, tmpdir, tools, timeout):
    """LibreOffice headless -> first-page PNG. Isolated profile avoids clashing with a live LO."""
    out_dir = os.path.join(tmpdir, 'png')
    os.makedirs(out_dir, exist_ok=True)
    profile = os.path.join(tmpdir, 'lo_profile').replace('\\', '/')
    uri = 'file:///' + profile.lstrip('/')
    r = _run([tools['soffice'], '--headless', '--norestore', '-env:UserInstallation=' + uri,
              '--convert-to', 'png', '--outdir', out_dir, docx_path], timeout)
    png = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.png')
    if not os.path.isfile(png):
        err = r.stderr.decode('utf-8', 'replace') if r.stderr else ''
        raise RuntimeError('soffice render failed: ' + err.strip()[:300])
    return png


def process_png(src_png, out_path, transparent=True, margin=12):
    """Crop tight to the ink, keep a margin, optionally make white pixels transparent."""
    from PIL import Image, ImageChops
    im = Image.open(src_png).convert('RGB')
    diff = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255)))
    bbox = diff.getbbox()
    if bbox:
        l = max(0, bbox[0] - margin)
        t = max(0, bbox[1] - margin)
        rgt = min(im.size[0], bbox[2] + margin)
        bot = min(im.size[1], bbox[3] + margin)
        im = im.crop((l, t, rgt, bot))
    if transparent:
        gray = im.convert('L')
        mask = gray.point(lambda p: 0 if p >= 250 else 255)  # 0 = white -> fully transparent
        rgba = im.convert('RGBA')
        rgba.putalpha(mask)
        rgba.save(out_path)
    else:
        im.save(out_path)
    return out_path


def render_png_matplotlib(latex, out_path, display=True, dpi=300):
    """Degraded renderer: matplotlib mathtext (subset of LaTeX; no cases/aligned)."""
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig = plt.figure(figsize=(0.01, 0.01))
    tex = '$' + latex + '$'
    if display:
        tex = '$\\displaystyle ' + latex + '$'
    fig.text(0, 0, tex, fontsize=14)
    fig.savefig(out_path, dpi=dpi, bbox_inches='tight', pad_inches=0.08, transparent=True)
    plt.close(fig)
    return out_path


def render_png(latex, out_path, tools, tmpdir, timeout, display=True, transparent=True):
    """Unified entry. Returns {'renderer': ...} on success, {'error':..., 'message':...} on failure.

    Fidelity order: pandoc+soffice (primary) -> matplotlib (only for infrastructure
    problems or missing tools; LaTeX syntax errors are NEVER masked by the fallback).
    """
    if pandoc_available(tools):
        # Always validate the LaTeX through pandoc first — a syntax error must fail loudly.
        docx_path = temp_docx_from_latex(latex, tmpdir, tools, timeout)
        if soffice_available(tools):
            try:
                png = soffice_png_from_docx(docx_path, tmpdir, tools, timeout)
                process_png(png, out_path, transparent)
                return {'renderer': 'pandoc+soffice'}
            except Exception as e:
                try:
                    render_png_matplotlib(latex, out_path, display)
                    return {'renderer': 'matplotlib (soffice unavailable: %s)' % str(e)[:120]}
                except Exception as e2:
                    return {'error': 'render_failed',
                            'message': 'soffice: %s | matplotlib: %s' % (str(e)[:200], str(e2)[:200])}
        # pandoc OK but no soffice -> matplotlib degraded (LaTeX already validated above)
        try:
            render_png_matplotlib(latex, out_path, display)
            return {'renderer': 'matplotlib (soffice not found)'}
        except Exception as e2:
            return {'error': 'render_failed', 'message': str(e2)[:300]}
    # no pandoc at all: matplotlib only
    try:
        render_png_matplotlib(latex, out_path, display)
        return {'renderer': 'matplotlib (pandoc not found — degraded)'}
    except Exception as e:
        return {'error': 'render_failed', 'message': 'pandoc missing and matplotlib failed: ' + str(e)[:200]}
