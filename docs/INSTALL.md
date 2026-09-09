# Installation

dsh-latex2office is a DSH Desktop / DeepSeek Harness bundle plugin. Install by
adding it to your profile manifest at `harness/profiles/web/package.json`.

## From GitHub

```jsonc
{
  "dependencies": { "dsh-latex2office": "github:hufangd-2019/dsh-latex2office" },
  "dsh": { "profile": { "bundles": ["@deepseek-ai/dsh-base", "...", "dsh-latex2office"] } }
}
```

Then run `pnpm install` in the profile directory and restart DSH Desktop.

## Host prerequisites

The plugin registers at boot with **no** external IO (it cannot break startup).
The actual conversion tools run only when a model calls them, and need these on
the host machine:

- **Python** >= 3.10 with `python-docx`, `python-pptx`, `lxml`, `Pillow`
- **pandoc** — LaTeX to native OMML equations
- **LibreOffice** — high-resolution image rendering for `.pptx` and preview

Each is discovered via cordis row config -> `PATH` -> common install locations.
A missing tool yields a clear JSON error from the affected tool only; it never
affects plugin loading or the rest of the harness.

## Verification

After restart, `harness.log` should contain:

```
[stdout] [latex2office] tools registered: latex2office_insert, latex2office_generate, latex2office_preview
```