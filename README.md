# handbook-tools

Shared build tooling for the St Margaret's College handbooks — the
[Members' Handbook](https://github.com/smcdn/Members-Handbook) and the
[College Leaders' Handbook](https://github.com/smcdn/Leaders-Handbook). Both are
MkDocs sites; this package is what turns them into a PDF with a running header
and a Word working copy for editors.

You almost certainly do not need this repo. Everything a content editor or a
publisher does is described in each handbook's own README.

## What's in it

| Module | What it does |
|---|---|
| `handbook.py` | Reads which handbook is being built (name, year, audience) from its `mkdocs.yml`. Everything else takes that as its input, which is why the same code builds both handbooks. |
| `hooks.py` | Generates the PDF's running header, and names the PDF, from `extra.dates.year`. |
| `docx_export.py` | Builds the Word working copy from the combined HTML the `to-pdf` plugin produces. |
| `docx_post.py` | The OOXML fix-ups pandoc cannot express: floating images, unbroken admonition boxes, a contents list that needs no updating, Track Changes switched on. |

## How a handbook uses it

The site declares the dependency in its `pyproject.toml`, and registers the
hooks through two shim files in its own repo:

```python
# hooks.py
from handbook_tools.hooks import on_config, on_files  # noqa: F401
```

```python
# docx_export.py
from handbook_tools.docx_export import on_post_build  # noqa: F401
```

They stay hooks rather than becoming a plugin for one reason: the DOCX is built
from the `to-pdf` plugin's output, so it must run *after* that plugin, and
MkDocs guarantees that by appending hooks after the plugins. As a plugin it
would depend on the order of the `plugins:` list instead.

## Changing it

There are no tests. The check that matters is building both handbooks with
`ENABLE_PDF_EXPORT=1 ENABLE_DOCX_EXPORT=1 uv run mkdocs build` and opening the
PDF and the DOCX.

Nothing here may name a particular handbook. Anything specific to one of them —
its title, who it is written for, which images the Word copy wraps text around —
belongs in the `extra.handbook` block of that site's `mkdocs.yml`, and reaches
the code through `handbook.Handbook`.

Each site pins this package by commit in its `uv.lock`, so a change lands only
when that site runs:

```
uv lock --upgrade-package handbook-tools
```
