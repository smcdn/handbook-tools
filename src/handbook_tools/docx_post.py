"""OOXML fix-ups applied to the .docx pandoc produces, for docx_export.py.

These are the things pandoc cannot express coming from HTML: text wrapping around
an image (it only ever emits inline drawings), keeping an admonition box whole
across a page break, a contents list that needs no updating on open, Track Changes
switched on, and the provenance stamp in the document properties.
"""
import re
import zipfile

EMU_PER_TWIP = 635


# pandoc writes <w:pStyle ... />, the template writes <w:pStyle .../>; match both.
_PSTYLE = re.compile(r'<w:pStyle w:val="([^"]*)"\s*/>')


def _paragraphs(doc):
    """(start, end, text) for every <w:p> in document order. w:p never nests."""
    return [(m.start(), m.end(), m.group()) for m in re.finditer(r"<w:p>.*?</w:p>", doc, re.S)]


def _drawing_para(doc, title):
    """The paragraph containing the drawing whose docPr title matches."""
    for start, end, text in _paragraphs(doc):
        if f'title="{title}"' in text:
            return start, end, text
    raise LookupError(f"no drawing titled {title!r}")


def _to_anchor(drawing, side, dist_left, dist_right):
    """Turn an inline drawing into a floating one that text wraps around."""
    inner = re.search(r"<wp:inline>(.*)</wp:inline>", drawing, re.S).group(1)
    extent = re.search(r"<wp:extent[^>]*/>", inner).group()
    effect = re.search(r"<wp:effectExtent[^>]*/>", inner).group()
    rest = inner.replace(extent, "", 1).replace(effect, "", 1)
    return (
        f'<wp:anchor distT="0" distB="91440" distL="{dist_left}" distR="{dist_right}"'
        ' simplePos="0" relativeHeight="2" behindDoc="0" locked="0"'
        ' layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:align>{side}</wp:align></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f"{extent}{effect}"
        '<wp:wrapSquare wrapText="bothSides"/>'
        f"{rest}</wp:anchor>"
    )


def float_image(doc, title, side="right"):
    """Float a standalone image so body text wraps beside it."""
    start, end, para = _drawing_para(doc, title)
    left, right = (182880, 0) if side == "right" else (0, 182880)
    floated = re.sub(r"<wp:inline>.*</wp:inline>", lambda m: _to_anchor(m.group(), side, left, right),
                     para, flags=re.S)
    return doc[:start] + floated + doc[end:]


def float_figure(doc, title, side="right"):
    """Float an image *and its caption* together, in a borderless one-cell table.

    A bare anchored image would leave the caption behind in the text flow, so the
    pair goes into a floating table - the way this is done by hand in Word.
    """
    start, end, para = _drawing_para(doc, title)
    caption = re.match(r'<w:p><w:pPr><w:pStyle w:val="ImageCaption" />.*?</w:p>',
                       doc[end:], re.S)
    body = para + (caption.group() if caption else "")
    end += caption.end() if caption else 0

    cx = int(re.search(r'<wp:extent cx="(\d+)"', para).group(1))
    width = cx // EMU_PER_TWIP + 120  # image width plus cell margins
    none = "".join(f'<w:{e} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                   for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    table = (
        "<w:tbl><w:tblPr>"
        f'<w:tblpPr w:leftFromText="181" w:rightFromText="181" w:topFromText="0"'
        f' w:bottomFromText="0" w:vertAnchor="text" w:horzAnchor="margin"'
        f' w:tblpXSpec="{side}" w:tblpY="1"/>'
        f'<w:tblW w:type="dxa" w:w="{width}"/>'
        f"<w:tblBorders>{none}</w:tblBorders>"
        '<w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0"'
        ' w:noHBand="1" w:noVBand="1" w:val="0000"/>'
        f'</w:tblPr><w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{width}"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr>'
        f"{body}</w:tc></w:tr></w:tbl>"
        "<w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"20\" w:lineRule=\"exact\"/></w:pPr></w:p>"
    )
    return doc[:start] + table + doc[end:]


def keep_boxes_whole(doc, prefix=("Admonition", "EditingNotice")):
    """Stop a note/warning box splitting across a page break.

    keepLines holds each paragraph together; keepNext holds the paragraphs of one
    box to each other. The last paragraph of a box gets no keepNext, so the box
    isn't glued to the body text that follows it.
    """
    paras = _paragraphs(doc)
    styles = [_PSTYLE.search(p[2]) for p in paras]
    styles = [m.group(1) if m else None for m in styles]
    out, prev_end = [], 0
    for i, (start, end, text) in enumerate(paras):
        style = styles[i]
        if style and style.startswith(prefix):
            last = i + 1 >= len(styles) or styles[i + 1] != style
            keep = "<w:keepLines/>" if last else "<w:keepNext/><w:keepLines/>"
            text = _PSTYLE.sub(lambda m: m.group() + keep, text, count=1)
        out.append(doc[prev_end:start] + text)
        prev_end = end
    return "".join(out) + doc[prev_end:]


def fill_contents(doc, levels=1):
    """Give the contents field a result, and stop Word offering to update it.

    pandoc marks the TOC field w:dirty="true", which is what makes Word ask "this
    document contains fields that may refer to other files" on opening. Filling the
    field in ourselves - chapter titles, no page numbers (the \\n switch), so
    nothing can go stale - means the list is right on open with nothing to update.

    `levels` is how deep this handbook's chapters sit: 1 where the nav is flat, 2
    where it groups chapters under sections. See Handbook.chapter_level.
    """
    headings = {f"Heading{n}": n for n in range(1, levels + 1)}
    entries = ""
    for _, _, text in _paragraphs(doc):
        m = _PSTYLE.search(text)
        if not m or m.group(1) not in headings:
            continue
        title = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", text))
        entries += (f'<w:p><w:pPr><w:pStyle w:val="TOC{headings[m.group(1)]}"/></w:pPr>'
                    f'<w:r><w:t xml:space="preserve">{title}</w:t></w:r></w:p>')
    field = (
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve">'
        f'TOC \\o &quot;1-{levels}&quot; \\h \\z \\u \\n'
        '</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        f'{entries}'
        '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )
    # a lambda, so the backslashes in the field instruction survive re.sub
    return re.sub(r'<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true".*?</w:p>',
                  lambda _: field, doc, count=1, flags=re.S)


def set_provenance(core, text):
    return core.replace("<cp:keywords></cp:keywords>",
                        f"<cp:keywords></cp:keywords><dc:description>{text}</dc:description>")


def enable_track_changes(settings):
    """Open with Track Changes already recording.

    pandoc drops <w:trackChanges/> from the reference document's settings, so it
    has to be put back here. Schema order puts it just before w:doNotTrackMoves.
    """
    return re.sub(r"<w:doNotTrackMoves\s*/>", "<w:trackChanges/><w:doNotTrackMoves/>",
                  settings, count=1)


def rewrite(path, transform_document, provenance=None, track_changes=True):
    zin = zipfile.ZipFile(path)
    parts = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zo:
        for item, blob in parts:
            if item.filename == "word/document.xml":
                blob = transform_document(blob.decode("utf8")).encode("utf8")
            elif item.filename == "word/settings.xml" and track_changes:
                blob = enable_track_changes(blob.decode("utf8")).encode("utf8")
            elif item.filename == "docProps/core.xml" and provenance:
                blob = set_provenance(blob.decode("utf8"), provenance).encode("utf8")
            zo.writestr(item, blob)
