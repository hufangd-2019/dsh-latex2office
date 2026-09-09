# Tool & engine reference

## Tools registered into the DSH tools registry

### `latex2office_insert`

| Arg | Type | Required | Notes |
|---|---|---|---|
| `file` | string | yes | existing .docx or .pptx (absolute or workspace-relative) |
| `formulas` | array | yes | one object per formula (below) |
| `backup` | boolean | no | default true; first modification writes name.bak.ext |
| `slide_index` | integer | no | pptx only: 1-based slide number |

Each `formulas` item:

| Key | Type | Notes |
|---|---|---|
| `latex` | string | required; a stray outer math wrapper is auto-stripped |
| `position` | object | `{mode: "end" \| "after_text" \| "before_text" \| "after_paragraph", text?, index?}` |
| `math_mode` | "omml" \| "image" | docx default omml; pptx default image |
| `number` | string | e.g. "3.1"; academic tab layout (centered + flush-right) |
| `display` | boolean | inline (false) vs display (true, default); docx only |

Response: `{ok:true, file, backup, inserted:[{ok, mode, number, verified}], touched}`.

### `latex2office_generate`

| Arg | Type | Required | Notes |
|---|---|---|---|
| `output` | string | yes | new .docx path |
| `title` | string | no | Heading 1 at the top |
| `formulas` | array | yes | as above (only latex/number/display matter) |

Response: `{ok:true, file, count}`.

### `latex2office_preview`

| Arg | Type | Required | Notes |
|---|---|---|---|
| `latex` | string | yes | a stray outer wrapper is auto-stripped |
| `display` | boolean | no | default true |

Response: `{ok:true, png, renderer}`. `renderer` is one of
`pandoc+soffice` (primary), `matplotlib (...)` (degraded when soffice/pandoc
is unavailable).

## Position modes

- `end` (default) — append at the document end.
- `after_text` — right after the first occurrence of `position.text` in the
  original document text.
- `before_text` — right before the first occurrence of `position.text`.
- `after_paragraph` — at the 1-based `position.index` paragraph.

Anchors are matched against the ORIGINAL document; a miss yields
`validation_error` and, because insertion is atomic, the file is left
byte-for-byte untouched.

## Error codes

All handled errors return `{ok:false, error:"<code>", message:"..."}`:

| code | cause |
|---|---|
| `missing_file` / `missing_output` | a required path arg was absent |
| `bad_formulas` | formulas is not a non-empty array |
| `file_not_found` | target file does not exist |
| `bad_file_type` | not .docx/.pptx |
| `file_locked` | another app holds the file (probe_writable) |
| `validation_error` | bad anchor text, missing position mode, empty latex |
| `latex_error` | pandoc rejected the LaTeX (syntax) |
| `render_failed` | soffice + matplotlib both failed |
| `python_not_found` / `pandoc_not_found` / `soffice_not_found` | a host tool is missing |
| `python_deps_missing` | python-docx/python-pptx/lxml/Pillow not installed |
| `engine_invalid_output` | engine stdout was not JSON |
| `engine_timeout` | subprocess exceeded the grace window |
| `tool_exception` | unexpected throw (defensive catch) |

## Engine JSON protocol

`engine/engine.py` reads ONE JSON request from stdin and writes ONE JSON
result to stdout, exit code 0 on both success and handled errors. It never
raises; the only non-zero exit is a fatal Python error.

```
{"action":"insert|generate|preview",
 "file": "...",
 "output": "...",
 "title": "...",
 "formulas":[...],
 "latex":"...",
 "display": true,
 "preview_dir":"...",
 "renderTimeoutSec": 90,
 "backup": true,
 "tools":{"python":"...","pandoc":"...","soffice":"..."}}
```

The host (lib/index.js) resolves `tools` via cordis config, then
`ctx.subprocess.resolveExecutable`, then a common-paths table; the engine
re-checks each with `shutil.which` plus the table and never raises on a miss.