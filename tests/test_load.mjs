/**
 * Isolated load test for dsh-latex2office — crash-safety verification.
 *
 * Simulates a minimal cordis context (tools registry + subprocess service) and:
 *   1. imports the plugin module (catches syntax/import errors)
 *   2. runs apply(ctx, config) (must not throw, registers exactly 3 tools)
 *   3. executes latex2office_generate end-to-end through a REAL spawned python
 *      engine subprocess (the same argv/stdin/stdout contract the harness uses)
 *   4. executes latex2office_preview end-to-end (needs pandoc + LibreOffice)
 *   5. executes latex2office_insert end-to-end on the generated sample
 *   6. verifies the atomic-failure path returns a clean JSON error (not a crash)
 *
 * Run with: node tests/test_load.mjs
 * Requirements: Python >=3.10 with python-docx/pandas-free stack
 *   (python-docx, python-pptx, lxml, Pillow), pandoc on PATH; LibreOffice for
 *   the preview step. Override discovery with env vars:
 *   L2O_TEST_PYTHON / L2O_TEST_PANDOC / L2O_TEST_SOFFICE (absolute paths).
 * All artifacts go into a fresh temp dir. This NEVER touches the real DSH
 * profile or the running desktop app.
 */
import { spawn, spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const PLUGIN = new URL('../lib/index.js', import.meta.url)
const REPO_ROOT = fileURLToPath(new URL('..', import.meta.url))
const ENGINE_PY = join(REPO_ROOT, 'engine', 'engine.py')
const WORKDIR = join(tmpdir(), 'dsh-latex2office-loadtest-' + Date.now().toString(36))
mkdirSync(WORKDIR, { recursive: true })

function whichSync(cmd) {
  try {
    const r = spawnSync(process.platform === 'win32' ? 'where' : 'which', [cmd], { encoding: 'utf-8', windowsHide: true })
    const first = (r.stdout || '').split(/\r?\n/)[0]
    return first && existsSync(first) ? first : ''
  } catch {
    return ''
  }
}
function findTool(envName, names, knownDirs) {
  if (process.env[envName] && existsSync(process.env[envName])) return process.env[envName]
  for (const n of names) {
    const p = whichSync(n)
    if (p) return p
  }
  for (const k of knownDirs) {
    if (existsSync(k)) return k
  }
  return ''
}
const PY = findTool('L2O_TEST_PYTHON', ['python', 'python3'], [])
const PANDOC = findTool('L2O_TEST_PANDOC', ['pandoc'], ['C:/Program Files/Pandoc/pandoc.exe', '/usr/bin/pandoc'])
const SOFFICE = findTool('L2O_TEST_SOFFICE', ['soffice', 'soffice.com', 'libreoffice'], [
  'C:/Program Files/LibreOffice/program/soffice.com',
  '/usr/bin/soffice',
])

function log(step, msg) {
  console.log('[loadtest] ' + step + ': ' + msg)
}
function fail(step, msg) {
  console.error('[loadtest] FAIL ' + step + ': ' + msg)
  try { rmSync(WORKDIR, { recursive: true, force: true }) } catch {}
  process.exit(1)
}
if (!PY) fail('env', 'no python found (set L2O_TEST_PYTHON to an absolute path)')
if (!PANDOC) fail('env', 'no pandoc found (set L2O_TEST_PANDOC to an absolute path)')

// ---- minimal mock of the ctx.subprocess service (DSH subprocess-local-compatible contract) ----
function mockSpawn(opts) {
  const child = spawn(opts.argv[0], opts.argv.slice(1), {
    cwd: opts.cwd,
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  })
  let out = ''
  let err = ''
  child.stdout.setEncoding('utf-8')
  child.stderr.setEncoding('utf-8')
  child.stdout.on('data', (d) => (out += d))
  child.stderr.on('data', (d) => (err += d))
  const cap = opts.stdio && opts.stdio.stdout && opts.stdio.stdout.maxBytes ? opts.stdio.stdout.maxBytes : 8 * 1024 * 1024
  let doneResolve
  const done = new Promise((resolve) => {
    doneResolve = resolve
  })
  let killed = false
  const timer = setTimeout(() => {
    killed = true
    child.kill()
  }, opts.graceMs || 300000)
  child.on('error', (e) => {
    clearTimeout(timer)
    doneResolve({ exitCode: -1, error: e.message })
  })
  child.on('close', (code) => {
    clearTimeout(timer)
    doneResolve({ exitCode: code, killed })
  })
  if (opts.stdio && opts.stdio.stdin && opts.stdio.stdin.data) {
    child.stdin.write(opts.stdio.stdin.data)
  }
  child.stdin.end()
  if (out.length > cap) out = out.slice(0, cap)
  return {
    done,
    collected: {
      stdout: { readFrom: () => ({ text: out }) },
      stderr: { readFrom: () => ({ text: err }) },
    },
  }
}
async function mockResolveExecutable(name) {
  const map = { python: PY, pandoc: PANDOC, soffice: SOFFICE }
  const p = map[name]
  if (p) return p
  throw new Error('not found: ' + name)
}

const registered = []
const ctx = {
  effect: (fn) => {
    fn()
  },
  tools: {
    register: (tool) => {
      registered.push(tool)
    },
  },
  subprocess: {
    spawn: mockSpawn,
    resolveExecutable: mockResolveExecutable,
  },
  sandboxPolicy: { workspaceRoot: WORKDIR },
}

// ---- 1. import the plugin ----
let plugin
try {
  const mod = await import(PLUGIN.href)
  plugin = mod.default
} catch (e) {
  fail('import', String(e))
}
log('import', 'ok (name=' + plugin.name + ', inject=' + JSON.stringify(plugin.inject) + ')')

// ---- 2. apply() ----
try {
  plugin.apply(ctx, {})
} catch (e) {
  fail('apply', 'apply() threw: ' + String(e))
}
if (registered.length !== 3) {
  fail('apply', 'expected 3 tools registered, got ' + registered.length)
}
const names = registered.map((t) => t.name).sort()
log('apply', 'registered: ' + names.join(', '))
const byName = Object.fromEntries(registered.map((t) => [t.name, t]))

// ---- 3. generate end-to-end (creates the sample used by step 5) ----
const sample = join(WORKDIR, 'sample.docx')
try {
  const res = await byName['latex2office_generate'].execute(
    {
      output: sample,
      title: 'loadtest sample',
      formulas: [{ latex: 'a^2 + b^2 = c^2', number: '0.1' }],
    },
    {}
  )
  if (res && res.ok && existsSync(res.file)) {
    log('generate-e2e', 'ok, file=' + res.file)
  } else {
    fail('generate-e2e', JSON.stringify(res).slice(0, 400))
  }
} catch (e) {
  fail('generate-e2e', 'threw (must never throw): ' + String(e))
}

// ---- 4. preview end-to-end (pandoc + LibreOffice main pipeline) ----
try {
  const res = await byName['latex2office_preview'].execute({ latex: 'e^{i\\pi} + 1 = 0', display: true }, {})
  if (res && res.ok && existsSync(res.png)) {
    log('preview-e2e', 'ok, png=' + res.png)
  } else {
    fail('preview-e2e', JSON.stringify(res).slice(0, 300))
  }
} catch (e) {
  fail('preview-e2e', 'threw (must never throw): ' + String(e))
}

// ---- 5. insert end-to-end into the generated sample ----
try {
  const res = await byName['latex2office_insert'].execute(
    {
      file: sample,
      formulas: [{ latex: '\\Gamma(n) = (n-1)!', position: { mode: 'end' }, number: '9.9' }],
      backup: true,
    },
    {}
  )
  if (res && res.ok && res.inserted && res.inserted.length === 1 && res.inserted[0].verified) {
    log('insert-e2e', 'ok, backup=' + res.backup)
  } else {
    fail('insert-e2e', JSON.stringify(res).slice(0, 400))
  }
} catch (e) {
  fail('insert-e2e', 'threw (must never throw): ' + String(e))
}

// ---- 6. failure path returns clean JSON, never throws ----
try {
  const res = await byName['latex2office_insert'].execute(
    {
      file: sample,
      formulas: [{ latex: 'E = mc^2', position: { mode: 'after_text', text: 'THIS TEXT DOES NOT EXIST 12345' } }],
    },
    {}
  )
  if (res && res.ok === false && res.error === 'validation_error') {
    log('failure-path', 'ok (clean JSON error: ' + res.message.slice(0, 60) + '...)')
  } else {
    fail('failure-path', JSON.stringify(res).slice(0, 300))
  }
} catch (e) {
  fail('failure-path', 'threw (must never throw): ' + String(e))
}

// cleanup
try {
  rmSync(WORKDIR, { recursive: true, force: true })
} catch {}

console.log('[loadtest] ALL PASS — plugin loads cleanly, tools execute end-to-end, failures are clean JSON.')
