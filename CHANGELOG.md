# Changelog

All notable changes to dsh-latex2office are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[SemVer](https://semver.org/).

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