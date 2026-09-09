// Structural check: the three tools must register object-rooted JSON Schema
// parameters (type/properties/required) — the shape the DSH model wire expects.
const mod = (await import('file:///F:/AImodel/dsh-latex2office-plugin/lib/index.js')).default
const registered = []
const ctx = {
  effect: (fn) => fn(),
  tools: { register: (t) => registered.push(t) },
  subprocess: { spawn: () => {}, resolveExecutable: async () => { throw new Error('n/a') } },
  sandboxPolicy: { workspaceRoot: 'F:\\AImodel\\iflow' },
}
mod.apply(ctx, {})
if (registered.length !== 3) {
  console.error('FAIL: expected 3 tools, got ' + registered.length)
  process.exit(1)
}
let bad = 0
for (const t of registered) {
  const p = t.parameters
  const ok = p && p.type === 'object' && p.properties && typeof p.properties === 'object' && Array.isArray(p.required)
  // no top-level reserved-keyword collision: every top-level key must be one of
  // the JSON Schema structural/annotation keywords we actually use
  const topKeys = Object.keys(p)
  const reservedNonString = topKeys.filter((k) => k !== 'type' && k !== 'properties' && k !== 'required')
  if (reservedNonString.length) {
    console.error(t.name + ': unexpected top-level keys: ' + reservedNonString.join(','))
    bad = 1
  }
  console.log((ok ? 'OK  ' : 'BAD ') + t.name + ' (properties: ' + Object.keys(p.properties).join(',') + '; required: [' + p.required.join(',') + '])')
  if (!ok) bad = 1
}
process.exit(bad ? 1 : 0)
