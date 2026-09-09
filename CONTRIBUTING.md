# Contributing

Thanks for helping improve dsh-latex2office.

## Layout

- `lib/index.js` — the cordis plugin entry point; registers the three model
  tools and speaks to the DSH subprocess service.
- `engine/*.py` — the conversion engine (`engine.py` is the stdin/stdout JSON
  entry point; `omml_docx.py` splices OMML into docx; `pptx_img.py` inserts
  images into pptx; `render.py` runs pandoc + LibreOffice).
- `tests/test_load.mjs` — the isolated load test (no real DSH profile needed).

## Development loop

```bash
node tests/test_load.mjs
```

This is the same test CI runs. It imports the plugin into a mock cordis
context, applies it, and drives every tool through real subprocesses plus the
math-delimiter stripping assertions. It never touches a live desktop profile.

Tool discovery can be pinned for a custom environment:

```bash
L2O_TEST_PYTHON=/opt/python3.12/bin/python \
L2O_TEST_PANDOC=/usr/bin/pandoc \
L2O_TEST_SOFFICE=/usr/bin/soffice \
node tests/test_load.mjs
```

## Engine protocol

The engine reads exactly one JSON request from stdin and writes exactly one
JSON result to stdout, exiting 0 on both success and handled errors. Host
paths arrive in `request.tools`; the engine double-checks with `shutil.which`
plus a common-paths table and never raises on a missing tool.

Handled errors use `{ok:false, error:'<code>', message:'<human text>'}`.
Adding a feature means adding its `error` code and message and, where
applicable, a test case in `tests/test_load.mjs`.

## Commit and release

- Keep commits small and self-contained; describe what and why.
- Version bumps go in `package.json` and `CHANGELOG.md` together, tagged as
  `vX.Y.Z`.
- A release tarball is produced with `npm pack`; attach it to the corresponding
  GitHub Release for users who prefer prebuilt installs.

## Reporting issues

Use the issue templates. Include the tool name, the exact LaTeX, and the error
code/message from the returned JSON.