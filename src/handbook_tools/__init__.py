"""Build tooling shared by the St Margaret's College handbooks.

Each handbook site registers these as MkDocs hooks, through one-line shim files
in its own repo (see the README). Nothing here knows which handbook it is
building: that comes from mkdocs.yml, via handbook.Handbook.
"""
