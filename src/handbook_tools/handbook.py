"""What this handbook is, read from mkdocs.yml.

Both build hooks need the same few facts - which handbook is being built, and
which year's edition - and neither should carry them as literals: the same
tooling builds the Members' Handbook and the College Leaders' Handbook. It all
comes from `site_name` and the `extra.handbook` / `extra.dates` blocks.
"""
import re
from xml.sax.saxutils import escape

# pandoc escapes these when it writes the XML attributes docx_post.py matches on.
XML_ESCAPES = {"'": "&#39;", '"': "&quot;"}


def xml(text):
    """`text` as it appears inside the OOXML pandoc produces."""
    return escape(text, XML_ESCAPES)


class Handbook:
    """The identity of the edition being built."""

    def __init__(self, config):
        extra = config["extra"]
        handbook = extra.get("handbook") or {}
        self.year = extra["dates"]["year"]
        # "St Margaret's College Members' Handbook" - the full name, as published.
        self.title = config["site_name"]
        # "Members' Handbook" - the short name, for headers and file names.
        self.short_title = handbook.get("short_name") or self.title
        # Which version is authoritative, and why - a clause, because it differs:
        # one handbook is read on the website, another is printed and handed out.
        self.authority = (handbook.get("authority")
                          or "the published handbook is the one that counts")
        # Alt text of the images the Word copy floats; see docx_post.py.
        self.float_figures = handbook.get("float_figures") or []
        self.float_images = handbook.get("float_images") or []
        # How deep a chapter sits: 1 where every chapter is a page in a flat nav,
        # 2 where the nav groups chapters under sections, so the chapter titles
        # come out as Heading 2. Drives the Word contents list, the running
        # header and where page breaks fall.
        self.chapter_level = int(handbook.get("chapter_level", 1))

    @property
    def file_base(self):
        """`2027-Members-Handbook` - the stem shared by the PDF and the DOCX."""
        words = re.sub(r"[^\w\s-]", "", self.short_title)  # apostrophes, commas
        slug = re.sub(r"[\s_]+", "-", words.strip())
        return f"{self.year}-{slug}"
