# Security policy

## Reporting a vulnerability

Please DO NOT open a public GitHub issue for security problems. Instead, email
or open a private security advisory on GitHub (Security tab > Report a
vulnerability), describing the issue and, if possible, a reproduction.

This project runs third-party conversion tools (pandoc, LibreOffice, Python)
through the DSH subprocess service. A formula string is passed to pandoc as an
argv element (never through a shell), so there is no shell-injection vector in
that path. If you believe you have found a way for crafted LaTeX to break out
of the conversion pipeline, include the exact input.