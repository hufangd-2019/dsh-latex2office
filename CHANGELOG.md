# Changelog

All notable changes to dsh-latex2office are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[SemVer](https://semver.org/).

## [0.1.2] - 2026-09-09

### Fixed
- **Tool registration schema (breaks every session on enable, 400 upstream).**
  `ctx.tools.register()` forwards `parameters` **verbatim** to the model API as
  the tool's JSON Schema. The 0.1.1 build passed a *flat property map*
  (`{ file: {...}, title: {...}, ... }`), so the optional `title` property leaked
  to the schema **root** and collided with the reserved JSON-Schema `title`
  keyword, which the upstream draft-04 metaschema requires to be a *string*
  (`'title' is not of type 'string'`) → `400 INVALID_REQUEST` on the very next
  turn of any session with the plugin enabled. All three tools (`latex2office_insert`,
  `latex2office_generate`, `latex2office_preview`) now register an object-rooted
  schema: `{ type: 'object', properties: { ... }, required: [ ... ] }`, with
  `required` lifted out of each property into a root `required` array. The
  `execute()` argument handling is unchanged (it already read flat `args`).
- Verified against the JSON-Schema metaschema: each tool's `parameters` now
  validates cleanly; the old flat map is reproduced as the 400 regression.
- `dsh.plugin.json` version synced to `0.1.2` (was stale at `0.1.0`).

## [0.1.1] - 2026-09-09

### Added
- Tolerance for stray math delimiters: a single outer `$$..$$`, `$..$`, `\[..\]`,
  or `\(..\)` wrapper is stripped before the formula reaches pandoc (which
  rejects a doubled wrapper). Conservative: a `$..$` pair is stripped only when
  the inner text contains no other `$`.
- GitHub Actions CI running the isolated load test and engine byte-compile on
  every push and pull request.
- Storefront screenshot manifest (`screenshots.json`) with rendered samples.

### Fixed
- `preview_dir` is now built with `path.join` instead of string concatenation,
  avoiding mixed-separator paths on Windows.

## [0.1.0] - 2026-09-09

### Added
- Three model tools: `latex2office_insert`, `latex2office_generate`,
  `latex2office_preview`.
- Native OMML equation insertion into `.docx` (pandoc + python-docx splice).
- High-resolution image insertion into `.pptx` (LibreOffice headless, isolated
  profile, tight crop + white-to-alpha).
- Atomic batch insertion with per-write readback verification.
- Academic tab-stop equation numbering.
- Automatic `.bak` backups and `probe_writable` lock detection.