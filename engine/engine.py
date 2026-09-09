#!/usr/bin/env python3
"""dsh-latex2office engine: reads ONE JSON request from stdin, prints ONE JSON result to stdout.

Actions:
  insert   - insert formulas into an existing .docx/.pptx (in place, atomic, verified)
  generate - create a new .docx from a list of formulas
  preview  - render one LaTeX formula to a PNG

Contract with the Node host (lib/index.js):
  - request:  { action, ..., tools: {python, pandoc, soffice}, renderTimeoutSec }
  - result:   { ok: bool, ... } printed as a single JSON object
  - exit code is 0 even for handled errors (the JSON carries the error);
    the Node side treats unparseable output as engine_invalid_output
"""
import json
import os
import shutil
import sys
import tempfile
import traceback

sys.stdin.reconfigure(encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8')

KNOWN_TOOL_PATHS = {
    'python': [r'C:\ProgramData\anaconda\python.exe', r'C:\ProgramData\anaconda3\python.exe'],
    'pandoc': [r'C:\Program Files\Pandoc\pandoc.exe', r'C:\Program Files (x86)\Pandoc\pandoc.exe'],
    'soffice': [r'C:\Program Files\LibreOffice\program\soffice.com',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.com'],
}


def fill_tools(tools):
    """Node host resolves tools; engine double-checks with shutil.which + known paths.

    Never raises — missing tools are reported per-use with clear errors.
    """
    out = dict(tools or {})
    for key, names in (('python', ('python',)), ('pandoc', ('pandoc',)), ('soffice', ('soffice', 'soffice.com'))):
        cur = out.get(key) or ''
        if cur and os.path.isfile(cur):
            continue
        for n in names:
            try:
                w = shutil.which(n)
            except Exception:
                w = None
            if w:
                out[key] = w
                break
        cur = out.get(key) or ''
        if not (cur and os.path.isfile(cur)):
            for p in KNOWN_TOOL_PATHS.get(key, []):
                if os.path.isfile(p):
                    out[key] = p
                    break
    return out


def fail(error, message, **extra):
    out = {'ok': False, 'error': error, 'message': message}
    out.update(extra)
    return out


def check_python_deps():
    missing = []
    for mod, pip_name in (('docx', 'python-docx'), ('pptx', 'python-pptx'), ('lxml', 'lxml'), ('PIL', 'Pillow')):
        try:
            __import__(mod)
        except Exception:
            missing.append(pip_name)
    return missing


def probe_writable(path):
    """Windows-friendly lock probe: opening r+b fails fast when another app holds the file."""
    try:
        f = open(path, 'r+b')
        f.close()
        return None
    except PermissionError:
        return 'file_locked'
    except OSError as e:
        return 'file_unreadable: ' + str(e)


def make_backup(file, enabled):
    """First-modification .bak; an existing .bak is never overwritten (keeps the earliest original)."""
    if not enabled:
        return ''
    base, ext = os.path.splitext(file)
    bak = base + '.bak' + ext
    if os.path.exists(bak):
        return bak
    try:
        shutil.copy2(file, bak)
        return bak
    except OSError:
        return ''


def main():
    try:
        raw = sys.stdin.buffer.read().decode('utf-8-sig')
        req = json.loads(raw)
    except Exception as e:
        print(json.dumps(fail('bad_request', 'stdin is not valid JSON: ' + str(e)), ensure_ascii=False))
        return

    action = req.get('action', '')
    tools = fill_tools(req.get('tools'))
    try:
        render_timeout = int(req.get('renderTimeoutSec') or 90)
    except Exception:
        render_timeout = 90

    missing = check_python_deps()
    if missing:
        print(json.dumps(fail('python_deps_missing', 'Missing Python packages: ' + ', '.join(missing)
                               + '. Install with: pip install ' + ' '.join(missing)), ensure_ascii=False))
        return

    tmpdir = tempfile.mkdtemp(prefix='latex2office_')
    try:
        if action == 'insert':
            file = req.get('file', '')
            if not file or not os.path.isfile(file):
                print(json.dumps(fail('file_not_found', 'Target file does not exist: ' + str(file)), ensure_ascii=False))
                return
            ext = os.path.splitext(file)[1].lower()
            if ext not in ('.docx', '.pptx'):
                print(json.dumps(fail('bad_file_type', 'Only .docx and .pptx are supported, got: ' + ext), ensure_ascii=False))
                return
            lock = probe_writable(file)
            if lock:
                print(json.dumps(fail(lock, 'The file is locked by another application (open in WPS / Word / '
                                       + 'LibreOffice?). Close it and retry: ' + file), ensure_ascii=False))
                return
            backup = make_backup(file, bool(req.get('backup', True)))
            if ext == '.docx':
                from omml_docx import insert_docx
                result = insert_docx(req, tools, tmpdir, render_timeout)
            else:
                from pptx_img import insert_pptx
                result = insert_pptx(req, tools, tmpdir, render_timeout)
            if backup:
                result.setdefault('backup', backup)
            if not result.get('ok'):
                result.setdefault('message', '')
            print(json.dumps(result, ensure_ascii=False))
            return

        if action == 'generate':
            output = req.get('output', '')
            if not output:
                print(json.dumps(fail('missing_output', 'output is required.'), ensure_ascii=False))
                return
            parent = os.path.dirname(os.path.abspath(output))
            if not os.path.isdir(parent):
                print(json.dumps(fail('output_dir_missing', 'Parent directory does not exist: ' + parent), ensure_ascii=False))
                return
            if os.path.exists(output) and not req.get('overwrite'):
                print(json.dumps(fail('output_exists', 'Output file already exists (pass overwrite:true to replace): '
                                       + output), ensure_ascii=False))
                return
            from omml_docx import generate_docx
            result = generate_docx(req, tools, tmpdir, render_timeout)
            print(json.dumps(result, ensure_ascii=False))
            return

        if action == 'preview':
            from render import render_png
            latex = (req.get('latex') or '').strip()
            if not latex:
                print(json.dumps(fail('missing_latex', 'latex is required.'), ensure_ascii=False))
                return
            out_png = os.path.join(tmpdir, 'preview.png')
            info = render_png(latex, out_png, tools, tmpdir, render_timeout,
                              display=bool(req.get('display', True)))
            if info.get('error'):
                print(json.dumps(fail(info['error'], info.get('message', '')), ensure_ascii=False))
                return
            target_dir = req.get('preview_dir')
            if target_dir:
                try:
                    os.makedirs(target_dir, exist_ok=True)
                    stable = os.path.join(target_dir, 'preview_' + os.urandom(6).hex() + '.png')
                except OSError:
                    stable = os.path.join(tempfile.gettempdir(), 'latex2office_preview_'
                                          + os.urandom(6).hex() + '.png')
            else:
                stable = os.path.join(tempfile.gettempdir(), 'latex2office_preview_'
                                      + os.urandom(6).hex() + '.png')
            shutil.copyfile(out_png, stable)
            print(json.dumps({'ok': True, 'png': stable, 'renderer': info.get('renderer')}, ensure_ascii=False))
            return

        print(json.dumps(fail('unknown_action', 'Unknown action: ' + str(action)), ensure_ascii=False))

    except Exception:
        print(json.dumps(fail('engine_exception', traceback.format_exc(limit=8)), ensure_ascii=False))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == '__main__':
    main()
