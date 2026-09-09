# dsh-latex2office

LaTeX → Office formula plugin for [DeepSeek Harness](https://www.deepseek.com) (DSH Desktop).

Converts LaTeX math into **native, editable OMML equations** for Word/WPS or **high-resolution rendered images** for `.pptx`, inserts them into existing `.docx` / `.pptx` files, and generates new `.docx` documents — with atomic batch insertion, academic tab-stop equation numbering, automatic `.bak` backups, and readback verification.

## Tools

| Tool | Purpose |
|---|---|
| `latex2office_insert` | Insert LaTeX formulas into an existing `.docx` / `.pptx` (native OMML or rendered images) |
| `latex2office_generate` | Create a new `.docx` from a list of formulas |
| `latex2office_preview` | Render one formula to a PNG preview before inserting |

Behavior highlights:

- **docx defaults to `math_mode=omml`**: produces native OMML equations, double-click editable in Word / WPS Writer; pptx defaults to `math_mode=image` (WPS Presentation has known issues rendering pptx OMML), override per formula with `omml`.
- Insert positions: `end` (default) / `after_text` / `before_text` / `after_paragraph`; docx supports `display:false` inline insertion (equation embedded in a text run).
- **Atomic batches**: any failing formula (bad LaTeX, missing anchor) leaves the file byte-for-byte untouched; anchors always match the original document text.
- **Numbering**: `number:"3.1"` renders the academic layout — centered equation + `(3.1)` flush right.
- First modification creates `name.bak.docx` / `name.bak.pptx` (never overwritten once present); every write is followed by a readback count.
- Complex structures map automatically: `aligned/align`→`m:eqArr`, `cases`/`\left(...\right)`→`m:d`, matrix environments→`m:m`.

## Host prerequisites (non-npm)

| Dependency | Used for | Notes |
|---|---|---|
| Python ≥ 3.10 | document engine | needs `python-docx`, `python-pptx`, `lxml`, `Pillow` |
| pandoc | LaTeX → OMML | extracts `m:oMathPara` from a temp docx and splices it in |
| LibreOffice | image rendering | `soffice --headless` with an isolated temp profile |

Resolution order: cordis row config → `PATH` → common install locations. A missing dependency only makes the affected tool return a clear JSON error — it never affects plugin loading or the rest of the harness.

## Install

### Option 1: DSH Desktop plugin market

Search `dsh-latex2office` in dshmarket and install with one click, then restart DSH Desktop.

### Option 2: Manual (GitHub dependency)

Edit the profile manifest (for Desktop: `harness/profiles/web/package.json`):

```jsonc
{
  "dependencies": {
    "dsh-latex2office": "github:hufangd-2019/dsh-latex2office"
  },
  "dsh": {
    "profile": {
      "bundles": ["dsh-latex2office"]
    }
  }
}
```

Then run `pnpm install` in the profile directory and restart DSH Desktop (the bundle layer is resolved at boot).

### Option 3: Local development (link)

Set `"dsh-latex2office": "link:<local dir>"` in `dependencies`, otherwise as above. Code changes take effect on restart.

### Optional config override

Attach a config to this plugin's row in the profile's own `cordis.patch.yml`:

```yaml
- id: latex2office-tools
  config:
    pythonPath: ''
    pandocPath: ''
    sofficePath: ''
    renderTimeoutSec: 90
```

## Usage (model-side calls)

```jsonc
// Generate a numbered-equation document
{ "tool": "latex2office_generate",
  "args": {
    "output": "D:/docs/equations.docx",
    "title": "Equations of motion",
    "formulas": [
      { "latex": "\\frac{\\partial u}{\\partial t} + u\\frac{\\partial u}{\\partial x} = -\\frac{1}{\\rho}\\frac{\\partial p}{\\partial x}", "number": "1.1" }
    ]
  } }

// Insert into an existing document after a text anchor
{ "tool": "latex2office_insert",
  "args": {
    "file": "D:/docs/report.docx",
    "formulas": [
      { "latex": "E = mc^2", "position": { "mode": "after_text", "text": "mass-energy relation:" }, "number": "2.1" }
    ]
  } }
```

## Uninstall / rollback

Remove the two added lines (dependency + bundle) from the profile manifest and restart. The plugin has **zero npm dependencies**, so no transitive packages are left behind.

## Crash-safety design

- `apply()` only registers tools and performs no external IO, so loading can never break harness boot.
- All heavy work (pandoc / Python / LibreOffice) runs through the DSH `ctx.subprocess` service with timeouts, stdout byte caps + spill protection, and cooperative `AbortSignal` cancellation; the harness main process stays responsive.
- Every tool `execute()` is fully try/caught and returns lossless JSON (`{ok:false, ...}`) instead of rejecting.
- LibreOffice rendering uses an isolated temporary profile and can never lock a document the user has open.

## Development

```bash
node tests/test_load.mjs
```

Isolated load test: simulates a cordis context and verifies import → apply (3 tool registrations, no throw) → generate/preview/insert end-to-end through real subprocesses → math-delimiter stripping (30 assertions on 10 input forms across 3 tools) → failure path returns clean JSON. Artifacts go to a temp directory and never touch a real DSH profile. Tool discovery can be overridden with `L2O_TEST_PYTHON` / `L2O_TEST_PANDOC` / `L2O_TEST_SOFFICE`.

## License

MIT