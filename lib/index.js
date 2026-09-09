/**
 * dsh-latex2office — LaTeX → Office formula tools (v0.1.1)
 *
 * Registers three model tools into the process-wide tools registry so every
 * session gets them:
 *   latex2office_insert   — insert LaTeX formulas into an EXISTING .docx/.pptx
 *                          (native OMML equations or rendered images;
 *                          positions end/after_text/before_text/after_paragraph;
 *                          atomic batch; numbered equations; .bak backup;
 *                          readback verification)
 *   latex2office_generate — create a NEW .docx from a list of formulas
 *   latex2office_preview  — render one LaTeX formula to a PNG preview image
 *
 * Crash-safety contract (hard requirement of the installation):
 *   - apply() only registers tools; it performs no external IO and cannot
 *     throw, so plugin loading can never break harness boot
 *   - all heavy work (pandoc / python engine / LibreOffice) is executed via
 *     the injected ctx.subprocess service with graceMs timeout, stdout byte
 *     caps and spill protection; the harness main process stays responsive
 *   - zero npm dependencies (pure ES module, no native modules)
 *   - every tool execute() is fully try/catch'ed and returns lossless JSON
 *     ({ok:false,...}) instead of rejecting
 *
 * Engine protocol: python engine/engine.py reads ONE JSON request from stdin
 * (request.tools carries resolved absolute paths) and prints ONE JSON result
 * object to stdout (exit code 0 in both success and handled-error cases).
 *
 * Tool path resolution order (per tool):
 *   1. cordis.patch.yml row config (pythonPath / pandocPath / sofficePath)
 *   2. ctx.subprocess.resolveExecutable (PATH lookup)
 *   3. common install locations table
 *   If none resolves, the affected tool returns a clear error at CALL time
 *   (never at load time).
 */
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ENGINE_PATH = fileURLToPath(new URL('../engine/engine.py', import.meta.url))

const KNOWN_PATHS = {
  python: [
    'C:\\ProgramData\\anaconda\\python.exe',
    'C:\\ProgramData\\anaconda3\\python.exe',
    'C:\\Python312\\python.exe',
  ],
  pandoc: [
    'C:\\Program Files\\Pandoc\\pandoc.exe',
    'C:\\Program Files (x86)\\Pandoc\\pandoc.exe',
  ],
  soffice: [
    'C:\\Program Files\\LibreOffice\\program\\soffice.com',
    'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com',
  ],
}

let subprocessRef = undefined
let cwdRef = '.'
let cfg = {}
let resolvedTools = undefined

/**
 * DSH tool results must be lossless JSON: objects, arrays, strings, numbers,
 * booleans, null — a single nested `undefined` makes the harness reject the
 * whole result ("value is not lossless JSON"). Strip undefined properties and
 * coerce non-finite numbers to null before returning anything from a tool.
 */
function pruneUndefined(value) {
  if (value === undefined) return null
  if (value === null) return null
  const t = typeof value
  if (t === 'string' || t === 'boolean') return value
  if (t === 'number') return Number.isFinite(value) ? value : null
  if (Array.isArray(value)) {
    const out = []
    for (let i = 0; i < value.length; i++) out.push(pruneUndefined(value[i]))
    return out
  }
  if (t === 'object') {
    const out = {}
    const keys = Object.keys(value)
    for (let i = 0; i < keys.length; i++) {
      const v = value[keys[i]]
      if (v === undefined) continue
      out[keys[i]] = pruneUndefined(v)
    }
    return out
  }
  return null
}

function jsonObjectOutput(properties) {
  return {
    schema: { type: 'object', properties: properties || {} },
    render: function (_args, value) {
      return [{ type: 'text', text: JSON.stringify(value) }]
    },
  }
}

const FORMULAS_DESC =
  'Array of formula specs. Each item: {latex: string (required; a stray $$/$ or \\[...\\] wrapper is auto-stripped), ' +
  'position: {mode: "end"|"after_text"|"before_text"|"after_paragraph", text?: string, index?: number} (default {mode:"end"}; ' +
  'after_text/before_text locate the first occurrence of text in the ORIGINAL document; after_paragraph uses a 1-based paragraph index), ' +
  'math_mode?: "omml"|"image" (default: omml for .docx, image for .pptx), ' +
  'number?: string (e.g. "3.1" — academic tab layout: centered formula + right-edge number; docx display formulas only), ' +
  'display?: boolean (default true = standalone centered paragraph; false = inline right after the position text — docx only)}'

const INSERT_OUTPUT = jsonObjectOutput({
  ok: { type: 'boolean' },
  file: { type: 'string' },
  fileType: { type: 'string' },
  inserted: { type: 'array' },
  backup: { type: 'string' },
  previewPng: { type: 'string' },
  error: { type: 'string' },
  message: { type: 'string' },
})

const GENERATE_OUTPUT = jsonObjectOutput({
  ok: { type: 'boolean' },
  file: { type: 'string' },
  inserted: { type: 'array' },
  error: { type: 'string' },
  message: { type: 'string' },
})

const PREVIEW_OUTPUT = jsonObjectOutput({
  ok: { type: 'boolean' },
  png: { type: 'string' },
  error: { type: 'string' },
  message: { type: 'string' },
})

async function resolveTools() {
  if (resolvedTools) return resolvedTools
  const out = { python: '', pandoc: '', soffice: '' }
  for (const name of ['python', 'pandoc', 'soffice']) {
    const configured = cfg[name + 'Path']
    if (configured && existsSync(configured)) {
      out[name] = configured
      continue
    }
    let found = ''
    try {
      found = await subprocessRef.resolveExecutable(name)
    } catch (e) {
      found = ''
    }
    if (found && existsSync(found)) {
      out[name] = found
      continue
    }
    for (const cand of KNOWN_PATHS[name] || []) {
      if (existsSync(cand)) {
        out[name] = cand
        break
      }
    }
  }
  resolvedTools = out
  return out
}

async function runEngine(request, timeoutSec, signal) {
  const tools = await resolveTools()
  if (!tools.python) {
    return {
      ok: false,
      error: 'python_not_found',
      message:
        'No usable Python found. Install Python >=3.10 with python-docx/python-pptx/lxml/Pillow, ' +
        'or set pythonPath in the dsh-latex2office cordis.patch.yml row config.',
    }
  }
  if (!tools.pandoc) {
    return {
      ok: false,
      error: 'pandoc_not_found',
      message: 'No usable pandoc found. Install pandoc, or set pandocPath in the cordis.patch.yml row config.',
    }
  }
  if (request.action !== 'generate' && !tools.soffice && usesRendering(request)) {
    return {
      ok: false,
      error: 'soffice_not_found',
      message:
        'No usable LibreOffice soffice found. Image/preview rendering needs it. ' +
        'Install LibreOffice, or set sofficePath in the cordis.patch.yml row config.',
    }
  }
  const handle = subprocessRef.spawn({
    argv: [tools.python, ENGINE_PATH],
    cwd: cwdRef,
    stdio: {
      stdin: { data: JSON.stringify({ ...request, tools: tools }) },
      stdout: { maxBytes: 8 * 1024 * 1024, spill: { maxBytes: 2 * 1024 * 1024 } },
      stderr: { maxBytes: 64 * 1024 },
    },
    graceMs: Math.max(30, timeoutSec) * 1000,
    signal: signal,
  })
  const outcome = await handle.done
  let text = ''
  if (handle.collected.stdout) text = handle.collected.stdout.readFrom(0).text
  let errText = ''
  if (handle.collected.stderr) errText = handle.collected.stderr.readFrom(0).text
  let parsed
  try {
    parsed = JSON.parse(text)
  } catch (e) {
    parsed = undefined
  }
  if (parsed === undefined || typeof parsed !== 'object') {
    return pruneUndefined({
      ok: false,
      error: 'engine_invalid_output',
      exitCode: outcome.exitCode,
      stdout: String(text).slice(0, 2000),
      stderr: String(errText).slice(0, 1000),
      message: 'The engine did not return valid JSON.',
    })
  }
  return pruneUndefined(parsed)
}

function usesRendering(request) {
  // pptx image mode and preview need LibreOffice; docx omml does not.
  if (request.action === 'preview') return true
  if (request.action === 'insert' || request.action === 'generate') {
    const list = request.formulas || []
    for (let i = 0; i < list.length; i++) {
      const f = list[i] || {}
      const mode = f.math_mode
      const isPptx = request.action === 'insert' && String(request.file || '').toLowerCase().endsWith('.pptx')
      if (isPptx && mode !== 'omml') return true
      if (!isPptx && mode === 'image') return true
    }
  }
  return false
}

function normalizeFormulas(raw) {
  if (!Array.isArray(raw)) return []
  const out = []
  for (let i = 0; i < raw.length; i++) {
    const f = raw[i] || {}
    const item = {
      latex: stripMathDelims(String(f.latex == null ? '' : f.latex).trim()),
      position: f.position && typeof f.position === 'object' ? f.position : { mode: 'end' },
      math_mode: f.math_mode === 'omml' || f.math_mode === 'image' ? f.math_mode : undefined,
      number: f.number == null ? undefined : String(f.number),
      display: f.display === undefined ? undefined : Boolean(f.display),
    }
    out.push(item)
  }
  return out
}

// Models occasionally paste formulas WITH their math delimiters ($x^2$, $$x^2$$,
// \[x^2\]). Pandoc rejects a doubled wrapper, so strip one matching outer pair.
// Conservative: a $...$ pair is stripped only when the inner text has no other $,
// so genuine content like "$a$ and $b$" or "\$5" is never touched.
function stripMathDelims(s) {
  if (s.length >= 4 && s.startsWith('$$') && s.endsWith('$$')) return s.slice(2, -2).trim() || s
  if (s.length >= 5 && s.startsWith('\\[') && s.endsWith('\\]')) return s.slice(2, -2).trim() || s
  if (s.length >= 5 && s.startsWith('\\(') && s.endsWith('\\)')) return s.slice(2, -2).trim() || s
  if (s.length >= 3 && s.startsWith('$') && s.endsWith('$')) {
    const inner = s.slice(1, -1)
    if (!inner.includes('$')) return inner.trim() || s
  }
  return s
}

const plugin = {
  name: 'latex2office-tools',
  inject: ['tools', 'subprocess', 'sandboxPolicy'],
  apply(ctx, config) {
    subprocessRef = ctx.subprocess
    cwdRef = (ctx.sandboxPolicy && ctx.sandboxPolicy.workspaceRoot) || '.'
    cfg = config && typeof config === 'object' ? config : {}

    ctx.effect(() =>
      ctx.tools.register({
        name: 'latex2office_insert',
        description:
          'Insert LaTeX formulas into an EXISTING Word/PowerPoint file (.docx/.pptx) and save it in place. ' +
          'docx: formulas become native OMML equations (editable in Word/WPS Writer; default). pptx: formulas become high-res rendered images by default (WPS Presentation has known issues displaying pptx OMML), pass math_mode:"omml" per formula to override. ' +
          'Positions: end (default, append at end), after_text / before_text (first occurrence of text in the ORIGINAL document), after_paragraph (1-based paragraph index). ' +
          'Batch is atomic: if any formula fails (bad LaTeX, missing anchor text), the file is left byte-for-byte untouched. ' +
          'A .bak copy is created on first modification (kept forever if it already exists; backup:false disables). ' +
          'Every insert is verified by re-opening the file. Auto formula numbering: number:"3.1" gives the academic layout (centered formula + right-edge "(3.1)"). ' +
          'For pptx pass slide_index (1-based); optional box {x,y,w} in inches centers by default. Use latex2office_preview first to eyeball a formula.',
        // Raw JSON Schema (register() passes `parameters` through verbatim to the
        // model API — NOT a flat property map; a top-level `title` key would collide
        // with the reserved JSON Schema `title` keyword and 400 upstream).
        parameters: {
          type: 'object',
          properties: {
            file: { type: 'string', description: 'Absolute path of the target .docx or .pptx file (must exist)' },
            formulas: { type: 'array', description: FORMULAS_DESC },
            slide_index: { type: 'number', description: 'pptx only: 1-based slide number to insert into (required for .pptx files)' },
            box: { type: 'object', description: 'pptx only: {x?, y?, w?} placement in inches; default centered, width = 60% of slide width' },
            backup: { type: 'boolean', description: 'Create/keep a .bak copy of the original before modification (default true)' },
          },
          required: ['file', 'formulas'],
        },
        output: INSERT_OUTPUT,
        timeoutMs: 300000,
        isConcurrencySafe: function () {
          return false
        },
        async execute(args, exec) {
          try {
            const file = String(args.file == null ? '' : args.file).trim()
            if (!file) return { ok: false, error: 'missing_file', message: 'file is required.' }
            const formulas = normalizeFormulas(args.formulas)
            if (formulas.length === 0) return { ok: false, error: 'missing_formulas', message: 'formulas (non-empty array) is required.' }
            for (let i = 0; i < formulas.length; i++) {
              if (!formulas[i].latex) {
                return { ok: false, error: 'missing_latex', message: 'formulas[' + i + '].latex is required.' }
              }
            }
            const request = {
              action: 'insert',
              file: file,
              formulas: formulas,
              slide_index: args.slide_index === undefined ? undefined : Math.floor(Number(args.slide_index)),
              box: args.box && typeof args.box === 'object' ? args.box : undefined,
              backup: args.backup === undefined ? true : Boolean(args.backup),
              renderTimeoutSec: Number(cfg.renderTimeoutSec) > 0 ? Number(cfg.renderTimeoutSec) : 90,
            }
            return await runEngine(request, 300, exec && exec.signal)
          } catch (e) {
            return pruneUndefined({ ok: false, error: 'tool_exception', message: e && e.message ? e.message : String(e) })
          }
        },
      })
    )

    ctx.effect(() =>
      ctx.tools.register({
        name: 'latex2office_generate',
        description:
          'Create a NEW .docx document from a list of LaTeX formulas (each becomes a native editable OMML equation; ' +
          'optional title as Heading 1; optional per-formula number for the academic tab layout). ' +
          'For inserting into an existing file use latex2office_insert instead.',
        parameters: {
          type: 'object',
          properties: {
            output: { type: 'string', description: 'Absolute path of the .docx file to create (parent directory must exist; existing file is overwritten only if overwrite:true)' },
            title: { type: 'string', description: 'Optional document title (Heading 1)' },
            formulas: {
              type: 'array',
              description:
                'Array of {latex: string (required; a stray $$/$ wrapper is auto-stripped), number?: string (e.g. "3.1"), display?: boolean (default true)}',
            },
            overwrite: { type: 'boolean', description: 'Allow overwriting an existing output file (default false)' },
          },
          required: ['output', 'formulas'],
        },
        output: GENERATE_OUTPUT,
        timeoutMs: 240000,
        isConcurrencySafe: function () {
          return false
        },
        async execute(args, exec) {
          try {
            const output = String(args.output == null ? '' : args.output).trim()
            if (!output) return { ok: false, error: 'missing_output', message: 'output is required.' }
            const formulas = normalizeFormulas(args.formulas)
            if (formulas.length === 0) return { ok: false, error: 'missing_formulas', message: 'formulas (non-empty array) is required.' }
            for (let i = 0; i < formulas.length; i++) {
              if (!formulas[i].latex) {
                return { ok: false, error: 'missing_latex', message: 'formulas[' + i + '].latex is required.' }
              }
            }
            const request = {
              action: 'generate',
              output: output,
              title: args.title === undefined ? undefined : String(args.title),
              formulas: formulas,
              overwrite: args.overwrite === undefined ? false : Boolean(args.overwrite),
              renderTimeoutSec: Number(cfg.renderTimeoutSec) > 0 ? Number(cfg.renderTimeoutSec) : 90,
            }
            return await runEngine(request, 240, exec && exec.signal)
          } catch (e) {
            return pruneUndefined({ ok: false, error: 'tool_exception', message: e && e.message ? e.message : String(e) })
          }
        },
      })
    )

    ctx.effect(() =>
      ctx.tools.register({
        name: 'latex2office_preview',
        description:
          'Render ONE LaTeX formula to a PNG image (white background cropped tight + 12px padding) and return its path — ' +
          'use to eyeball what a formula will look like before inserting it. Renders via the same pipeline as insert (pandoc OMML + LibreOffice), ' +
          'so what you see matches what Word/WPS show.',
        parameters: {
          type: 'object',
          properties: {
            latex: { type: 'string', description: 'LaTeX code; a stray $$/$ or \\[...\\] wrapper is auto-stripped' },
            display: { type: 'boolean', description: 'Render as display style (default true) vs inline style' },
          },
          required: ['latex'],
        },
        output: PREVIEW_OUTPUT,
        timeoutMs: 120000,
        isConcurrencySafe: function () {
          return true
        },
        async execute(args, exec) {
          try {
            const latex = stripMathDelims(String(args.latex == null ? '' : args.latex).trim())
            if (!latex) return { ok: false, error: 'missing_latex', message: 'latex is required.' }
            const request = {
              action: 'preview',
              latex: latex,
              display: args.display === undefined ? true : Boolean(args.display),
              preview_dir:
                cwdRef && cwdRef !== '.'
                  ? join(String(cwdRef).replace(/[\\/]+$/, ''), 'latex2office-previews')
                  : undefined,
              renderTimeoutSec: Number(cfg.renderTimeoutSec) > 0 ? Number(cfg.renderTimeoutSec) : 90,
            }
            return await runEngine(request, 120, exec && exec.signal)
          } catch (e) {
            return pruneUndefined({ ok: false, error: 'tool_exception', message: e && e.message ? e.message : String(e) })
          }
        },
      })
    )

    const msg =
      '[latex2office] tools registered: latex2office_insert, latex2office_generate, latex2office_preview (engine: ' + ENGINE_PATH + ')'
    // double-write: console.log lands in harness.log as [stdout] (same as dsh-deerflow),
    // ctx.logger.info reaches the in-app logger with structured metadata.
    console.log(msg)
    if (ctx && ctx.logger && typeof ctx.logger.info === 'function') {
      ctx.logger.info(msg)
    }
  },
}

export default plugin
