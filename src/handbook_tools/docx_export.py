"""MkDocs hook: build the Word working copy of the handbook.

Registered via `hooks:` in mkdocs.yml, and gated behind ENABLE_DOCX_EXPORT=1, so
routine builds and `mkdocs serve` never pay for it. MkDocs appends hooks to the
plugin collection, so this on_post_build runs after the to-pdf plugin's and can
read the combined HTML that plugin writes (its `html_path` option) - the same
document it hands to WeasyPrint, so the DOCX gets the same content, in the same
order, with the macros already expanded. ENABLE_PDF_EXPORT=1 is therefore
required as well.

The result is a working copy for drafting the next edition: Track Changes on,
Word styles instead of the print CSS, and prominent notices telling the editor
what it is. It is written to docx/ in the repo, never into site_dir, so it is
never served or deployed. Only the newest copy is kept.
"""
import glob
import os
import re
import subprocess
import tempfile
import zipfile
from datetime import date

from bs4 import BeautifulSoup

from . import docx_post
from .handbook import Handbook, xml

ENV_FLAG = "ENABLE_DOCX_EXPORT"
COMBINED = "combined.html"  # must match `html_path` on the to-pdf plugin
OUT_DIR = "docx"
MARGINS = (1134, 1134, 1134, 1134)  # 20 mm all round, in twips
PAGE = (11906, 16838)  # A4 portrait, in twips

# ------------------------------------------------------------------ text ----
WARNING = "Turn on Track Changes before you edit  —  Review tab → Track Changes"
NOTICE_TITLE = "Working copy — for editing only"
DATES_LEAD = "Leave dates and years as they are. "


def notice_body(book):
    return (
        "Track Changes is already switched on in this file; please check it is still "
        f"on before you type. This Word version of the {book.short_title} is here so "
        "that changes can be drafted and marked up. It is not the handbook itself — "
        f"the version on the College website is the one {book.audience} read, and it "
        "is the one that counts. The layout here only comes close to the published "
        "handbook: spacing, the position of pictures, table widths and page breaks "
        "will all differ, and there is no need to correct them. This copy was made "
        f"on {date.today():%-d %B %Y}."
    )


def dates_body(year):
    return (
        "Everything that changes from one intake to the next — the year itself, and "
        "the residential dates that go with it — is filled in automatically when the "
        f"handbook is published. Where this copy says {year}, the published handbook "
        "will say the right year. Please spend your time on the wording instead."
    )


# ---------------------------------------------------------------- styles ----
ADMONITIONS = {  # Material admonition type -> (border, fill)
    "Note": ("016A96", "E8F2F7"), "Info": ("00A0C6", "E4F5FA"),
    "Tip": ("00A88F", "E4F7F4"), "Summary": ("0091EA", "E6F4FE"),
    "Warning": ("E08600", "FFF4E2"), "Danger": ("C62828", "FDEAEA"),
    "Example": ("7B1FA2", "F5EAF7"), "Quote": ("757575", "F0F0F0"),
}


def box_style(style_id, name, border, fill):
    """A bordered, shaded paragraph style - one admonition box."""
    return f'''<w:style w:type="paragraph" w:styleId="{style_id}">
    <w:name w:val="{name}"/><w:basedOn w:val="BodyText"/><w:qFormat/>
    <w:pPr><w:pBdr>
      <w:top w:val="single" w:sz="4" w:space="6" w:color="{border}"/>
      <w:left w:val="single" w:sz="24" w:space="6" w:color="{border}"/>
      <w:bottom w:val="single" w:sz="4" w:space="6" w:color="{border}"/>
      <w:right w:val="single" w:sz="4" w:space="6" w:color="{border}"/>
    </w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>
    <w:spacing w:before="120" w:after="120"/><w:ind w:left="142" w:right="142"/>
    </w:pPr></w:style>'''


WARNING_STYLE = '''<w:style w:type="paragraph" w:styleId="TrackChangesWarning">
    <w:name w:val="Track Changes Warning"/><w:basedOn w:val="BodyText"/><w:qFormat/>
    <w:pPr><w:pBdr>
      <w:top w:val="single" w:sz="18" w:space="6" w:color="7F0000"/>
      <w:left w:val="single" w:sz="18" w:space="6" w:color="7F0000"/>
      <w:bottom w:val="single" w:sz="18" w:space="6" w:color="7F0000"/>
      <w:right w:val="single" w:sz="18" w:space="6" w:color="7F0000"/>
    </w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="B71C1C"/>
    <w:spacing w:before="120" w:after="120"/><w:ind w:left="142" w:right="142"/>
    <w:jc w:val="center"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/>
    <w:color w:val="FFFFFF"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
    </w:style>'''

def toc_style(level):
    """One contents-list level. w:name must be Word's built-in "toc N", so Word
    reuses the style if the reader ever rebuilds the list."""
    indent = f'<w:ind w:left="{(level - 1) * 220}"/>' if level > 1 else ""
    return (f'<w:style w:type="paragraph" w:styleId="TOC{level}">'
            f'<w:name w:val="toc {level}"/><w:basedOn w:val="Normal"/>'
            '<w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
            f'<w:pPr><w:spacing w:after="100"/>{indent}</w:pPr></w:style>')

TABLE_STYLE = '''<w:style w:type="table" w:default="1" w:styleId="Table">
    <w:name w:val="Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="99"/>
    <w:tblPr><w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
    </w:tblBorders><w:tblCellMar>
      <w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>
      <w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
    <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="016A96"/></w:tcPr>
    </w:tblStylePr></w:style>'''

def extra_styles(levels):
    """The styles added to pandoc's reference document."""
    return "".join(
        box_style(f"Admonition{k}", f"Admonition {k}", b, f)
        for k, (b, f) in ADMONITIONS.items()
    ) + box_style("EditingNotice", "Editing Notice", "B71C1C", "FDECEA") \
      + WARNING_STYLE + "".join(toc_style(n) for n in range(1, levels + 1))

# ---------------------------------------------- reference doc and template ----
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
SMALL = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/>'
         '<w:color w:val="808080"/></w:rPr>')


def field(instr):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t></w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def build_reference(pandoc, dst, book):
    """pandoc's own reference document, restyled for the handbook.

    Generated at build time rather than committed, so there is no binary in the
    repo to keep in step with the styles above.
    """
    default = os.path.join(os.path.dirname(dst), "reference-default.docx")
    with open(default, "wb") as f:
        subprocess.run([pandoc, "--print-default-data-file", "reference.docx"],
                       stdout=f, check=True)

    right_tab = PAGE[0] - MARGINS[1] - MARGINS[3]
    header_left = f"WORKING COPY &#8212; {book.year} {xml(book.short_title)}"
    header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<w:hdr {NS}><w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="{right_tab}"/>'
              f'</w:tabs></w:pPr><w:r>{SMALL}<w:t xml:space="preserve">{header_left}'
              f'</w:t></w:r><w:r>{SMALL}<w:tab/></w:r>'
              + field(f'STYLEREF "Heading {book.chapter_level}" \\* MERGEFORMAT')
              + '</w:p></w:hdr>')
    footer = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<w:ftr {NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
              f'<w:r>{SMALL}<w:t xml:space="preserve">~ </w:t></w:r>'
              + field('PAGE \\* MERGEFORMAT')
              + f'<w:r>{SMALL}<w:t xml:space="preserve"> ~</w:t></w:r></w:p></w:ftr>')

    zin = zipfile.ZipFile(default)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = zin.read(item.filename).decode("utf8")
            if item.filename == "word/_rels/document.xml.rels":
                data = data.replace("</Relationships>",
                    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/header" Target="header1.xml"/>'
                    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
                    "</Relationships>")
            elif item.filename == "[Content_Types].xml":
                data = data.replace("</Types>",
                    '<Override PartName="/word/header1.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.wordprocessingml.header+xml"/>'
                    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
            elif item.filename == "word/settings.xml":
                data = data.replace('<w:themeFontLang w:val="en-US" />',
                                    '<w:themeFontLang w:val="en-NZ" />')
            elif item.filename == "word/styles.xml":
                data = data.replace('<w:lang w:val="en-US"', '<w:lang w:val="en-NZ"')
                chapters = "|".join(f"Heading{n}"
                                    for n in range(1, book.chapter_level + 1))
                data = re.sub(rf'(<w:style [^>]*w:styleId="(?:{chapters})">.*?<w:pPr>)',
                              r'\1<w:pageBreakBefore/>', data, flags=re.S)
                data = re.sub(r'(<w:style [^>]*w:styleId="(?:Heading1|Heading2)">.*?<w:pPr>)',
                              r'\1<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="2"'
                              r' w:color="016A96"/></w:pBdr>', data, flags=re.S)
                data = re.sub(r'<w:style [^>]*w:styleId="Table"[^>]*>.*?</w:style>',
                              TABLE_STYLE, data, flags=re.S)
                data = re.sub(r'<w:rFonts w:asciiTheme="minorHAnsi"[^/]*/>',
                              '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>', data)
                data = re.sub(r'<w:color w:val="0F4761"[^/]*/>',
                              '<w:color w:val="016A96"/>', data)
                data = data.replace(
                    "</w:styles>", extra_styles(book.chapter_level) + "</w:styles>")
            zo.writestr(item, data)
        zo.writestr("word/header1.xml", header)
        zo.writestr("word/footer1.xml", footer)


def messages_ooxml(book):
    """The banner and notices as raw OOXML.

    They go into the template rather than the body so that they sit directly under
    the document title, above the table of contents.
    """
    def para(style, runs):
        body = "".join(f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
                       for rpr, text in runs)
        return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{body}</w:p>'

    bold = "<w:rPr><w:b/></w:rPr>"
    return (para("TrackChangesWarning", [("", WARNING)])
            + para("EditingNotice", [(bold, NOTICE_TITLE)])
            + para("EditingNotice", [("", notice_body(book))])
            + para("EditingNotice", [(bold, DATES_LEAD), ("", dates_body(book.year))]))


def build_template(pandoc, dst, book):
    """pandoc's docx template, carrying the page setup and the notices."""
    top, right, bottom, left = MARGINS
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>'
            '<w:footerReference w:type="default" r:id="rId10"/>'
            f'<w:pgSz w:w="{PAGE[0]}" w:h="{PAGE[1]}"/>'
            f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}"'
            ' w:header="567" w:footer="567" w:gutter="0"/><w:titlePg/>'
            '<w:footnotePr><w:numRestart w:val="eachSect"/></w:footnotePr></w:sectPr>')
    tpl = subprocess.run([pandoc, "-D", "docx"], capture_output=True, text=True,
                         check=True).stdout
    # pandoc does not carry the reference doc's sectPr (page size, margins, and the
    # header/footer references) into its output, so page setup lives here instead.
    tpl = tpl.replace("$sectpr$", sect)
    tpl = tpl.replace("$if(toc)$", messages_ooxml(book) + "\n$if(toc)$", 1)
    with open(dst, "w", encoding="utf8") as f:
        f.write(tpl)


# ------------------------------------------------------------- transform ----
def prepare(path):
    """The combined HTML, adapted for Word."""
    with open(path, encoding="utf8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    # Twemoji <img> (:material-check:/:material-close:) would need rsvg-convert to
    # survive the conversion; substitute the characters they stand for.
    for img in soup.select("img.converted-twemoji"):
        classes = (img.get("class") or []) + (img.parent.get("class") or [])
        img.replace_with(soup.new_string("✓" if "green" in classes else "✗"))

    # Admonitions become the Word styles carried by the reference document. The
    # title is bolded rather than given a style of its own, so the box stays whole.
    for div in soup.select("div.admonition"):
        kind = next((c for c in div.get("class", []) if c != "admonition"), "note")
        div["custom-style"] = f"Admonition {kind.title()}"
        title = div.select_one(".admonition-title")
        if title:
            strong = soup.new_tag("strong")
            strong.extend(title.contents)
            title.clear()
            title.append(strong)

    # Chapter numbers are baked into the PDF's headings; in a document being
    # reordered and rewritten they would only go stale.
    for span in soup.select("span.pdf-order"):
        span.decompose()

    # The PDF's cover page and static contents give way to the document title and
    # a Word contents list.
    for element_id in ("doc-cover", "doc-toc"):
        element = soup.find(id=element_id)
        if element:
            element.decompose()

    for tag in soup.select("link[rel=stylesheet], style, script"):
        tag.decompose()
    return soup


def provenance(repo, book):
    sha = subprocess.run(["git", "-C", repo, "rev-parse", "--short", "HEAD"],
                         capture_output=True, text=True).stdout.strip()
    source = f" (source commit {sha})" if sha else ""
    return (f"Generated from the {book.year} {book.short_title}{source} on "
            f"{date.today():%Y-%m-%d}. Working copy - not for publication.")


def on_post_build(config, **kwargs):
    if os.environ.get(ENV_FLAG) != "1":
        return

    combined = os.path.join(config["site_dir"], COMBINED)
    if not os.path.exists(combined):
        raise RuntimeError(
            f"{combined} is missing. The DOCX is built from the HTML the to-pdf "
            f"plugin writes, so ENABLE_PDF_EXPORT=1 is needed as well as {ENV_FLAG}=1."
        )

    import pypandoc
    pandoc = pypandoc.get_pandoc_path()

    book = Handbook(config)
    repo = os.path.dirname(config["config_file_path"])
    out_dir = os.path.join(repo, OUT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(
        out_dir, f"{book.file_base}-for-editing-{date.today():%Y-%m-%d}.docx")

    with tempfile.TemporaryDirectory() as work:
        reference = os.path.join(work, "reference.docx")
        template = os.path.join(work, "template.xml")
        build_reference(pandoc, reference, book)
        build_template(pandoc, template, book)
        subprocess.run(
            [pandoc, "-f", "html", "-t", "docx", "--standalone",
             "--toc", f"--toc-depth={book.chapter_level}",
             "--reference-doc", reference, "--template", template,
             "-M", "lang=en-NZ", "-o", out, "-"],
            input=str(prepare(combined)).encode(), check=True)

    def transform(doc):
        # Wrapping, which pandoc cannot express: it only emits inline images.
        for title in book.float_figures:
            doc = docx_post.float_figure(doc, xml(title), side="right")
        for title in book.float_images:
            doc = docx_post.float_image(doc, xml(title), side="right")
        doc = docx_post.fill_contents(doc, book.chapter_level)
        return docx_post.keep_boxes_whole(
            doc, prefix=("Admonition", "EditingNotice", "TrackChangesWarning"))

    docx_post.rewrite(out, transform, provenance=provenance(repo, book))

    # The intermediate is not part of the site.
    os.remove(combined)

    # Only the newest working copy is kept. Swept after the new file is complete,
    # so a failed build leaves the previous copy in place.
    for stale in sorted(glob.glob(os.path.join(out_dir, "*-for-editing-*.docx"))):
        if os.path.abspath(stale) != os.path.abspath(out):
            os.remove(stale)
            print(f"INFO    -  Removed superseded working copy "
                  f'"{os.path.basename(stale)}".')

    print(f'INFO    -  Output a DOCX working copy to "{out}".')
