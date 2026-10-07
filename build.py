#!/usr/bin/env python3
"""Builds the Livro website in every language.

    python3 build.py

Sources                              Output
  templates/index.html + locales/      index.html,   tr/index.html,   de/index.html
  templates/legal.html + legal/<lang>/ terms.html,   tr/terms.html,   de/terms.html
                                       privacy.html, tr/privacy.html, de/privacy.html

English lives at the root; every other language in a folder named after it.
Run this after changing a template, a locale file or a Word file, then commit
the generated pages too (GitHub Pages serves them as they are).

Legal pages are read from Word files, keeping what the app's DocxDocument
keeps: the title, numbered sections, sub-sections, paragraphs, bulleted and
numbered lists, bold text, line breaks and tables. Bracketed placeholders such
as [SUPPORT EMAIL] are highlighted so unfinished details are easy to spot.
Everything from a paragraph starting with "⚠️" onwards is the author's own
checklist and is left out.
"""

import html
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parent
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# The first language is the default and is served from the site root.
LANGUAGES = ["en", "tr", "de"]

# Set to the published address (e.g. "https://livro.app/") to emit absolute
# hreflang links, which search engines prefer. Relative links work meanwhile.
SITE_URL = ""

# Shown in every page footer.
CONTACT_EMAIL = "hellolivro@gmail.com"

PAGES = [
    # output file, template, legal source (None for the home page), locale key for its title
    ("index.html", "index.html", None, None),
    ("terms.html", "legal.html", "TermsOfService.docx", "legal.terms"),
    ("privacy.html", "legal.html", "PrivacyPolicy.docx", "legal.privacy"),
]


# MARK: - Word → HTML

def numbering_formats(archive):
    """numId → {level: format}, e.g. "bullet" or "decimal"."""
    try:
        root = ElementTree.fromstring(archive.read("word/numbering.xml"))
    except KeyError:
        return {}
    abstract = {}
    for node in root.iter(f"{W}abstractNum"):
        levels = {}
        for lvl in node.iter(f"{W}lvl"):
            fmt = lvl.find(f"{W}numFmt")
            levels[int(lvl.get(f"{W}ilvl"))] = fmt.get(f"{W}val") if fmt is not None else "bullet"
        abstract[node.get(f"{W}abstractNumId")] = levels
    formats = {}
    for num in root.iter(f"{W}num"):
        ref = num.find(f"{W}abstractNumId")
        if ref is not None:
            formats[num.get(f"{W}numId")] = abstract.get(ref.get(f"{W}val"), {})
    return formats


def placeholders(text):
    return re.sub(r"(\[[A-Z0-9 /.,'-]+\])", r'<mark class="todo">\1</mark>', text)


def inline(paragraph):
    """The paragraph's text as HTML, with bold runs and line breaks."""
    parts = []
    for run in paragraph.iter(f"{W}r"):
        props = run.find(f"{W}rPr")
        weight = props.find(f"{W}b") if props is not None else None
        bold = weight is not None and weight.get(f"{W}val") not in ("0", "false")
        text = ""
        for child in run:
            if child.tag == f"{W}t":
                text += html.escape(child.text or "")
            elif child.tag in (f"{W}br", f"{W}cr"):
                text += "<br>"
            elif child.tag == f"{W}tab":
                text += " "
        if text:
            parts.append((bold, text))
    # Merge neighbouring runs with the same weight so <strong> isn't split.
    merged = []
    for bold, text in parts:
        if merged and merged[-1][0] == bold:
            merged[-1] = (bold, merged[-1][1] + text)
        else:
            merged.append((bold, text))
    out = "".join(f"<strong>{t}</strong>" if b and t.strip() else t for b, t in merged)
    return placeholders(out.strip())


def plain(paragraph):
    return "".join(t.text or "" for t in paragraph.iter(f"{W}t")).strip()


def convert(docx):
    """(title, body blocks, [(anchor, heading)]) for a Word file."""
    archive = zipfile.ZipFile(docx)
    formats = numbering_formats(archive)
    body = ElementTree.fromstring(archive.read("word/document.xml")).find(f"{W}body")

    title, blocks, toc = None, [], []
    open_lists = []  # stack of "ul"/"ol"

    def close_lists(to=0):
        while len(open_lists) > to:
            blocks.append(f"</{open_lists.pop()}>")

    for node in body:
        if node.tag == f"{W}tbl":
            close_lists()
            rows = []
            for i, row in enumerate(node.iter(f"{W}tr")):
                tag = "th" if i == 0 else "td"
                cells = [" ".join(inline(p) for p in cell.iter(f"{W}p")) for cell in row.iter(f"{W}tc")]
                rows.append("<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</tr>")
            blocks.append('<div class="table-wrap"><table>' + "".join(rows) + "</table></div>")
            continue
        if node.tag != f"{W}p":
            continue

        text = plain(node)
        if text.startswith("⚠️"):
            break
        if not text:
            continue

        props = node.find(f"{W}pPr")
        outline = props.find(f"{W}outlineLvl") if props is not None else None
        num = props.find(f"{W}numPr") if props is not None else None

        if outline is not None:
            close_lists()
            # The drafts' outline levels aren't consistent, so: the first
            # heading is the title, numbered headings ("4. …") are sections,
            # the rest sub-sections.
            if title is None:
                title = text
                continue
            number = re.match(r"^([0-9]+)[.)]\s", text)
            if number:
                # Anchors use the section number, so #section-11 is the same
                # section in every language.
                anchor = f"section-{number.group(1)}"
                toc.append((anchor, text))
                blocks.append(f'<h2 id="{anchor}">{placeholders(html.escape(text))}</h2>')
            else:
                blocks.append(f"<h3>{placeholders(html.escape(text))}</h3>")
        elif num is not None:
            ilvl = num.find(f"{W}ilvl")
            level = int(ilvl.get(f"{W}val")) if ilvl is not None else 0
            num_id = num.find(f"{W}numId").get(f"{W}val")
            kind = "ul" if formats.get(num_id, {}).get(level, "bullet") in ("bullet", "none") else "ol"
            close_lists(level + 1)
            while len(open_lists) <= level:
                blocks.append(f"<{kind}>")
                open_lists.append(kind)
            blocks.append(f"<li>{inline(node)}</li>")
        else:
            close_lists()
            blocks.append(f"<p>{inline(node)}</p>")

    close_lists()
    # "Livro Privacy Policy — Draft" → "Livro Privacy Policy" (in any language)
    title = re.sub(r"\s*[—–-]\s*(Draft|Taslak|Entwurf)$", "", title or "")
    return title, blocks, toc


# MARK: - Pages

def folder(lang):
    return "" if lang == LANGUAGES[0] else f"{lang}/"


def lookup(strings, key):
    value = strings
    for part in key.split("."):
        value = value[part]
    return value


def fill(template, values, strings):
    """Replaces {{name}} with values[name], or with the locale string at that path."""
    def replace(match):
        key = match.group(1)
        if key in values:
            return values[key]
        return lookup(strings, key)
    return re.sub(r"\{\{([a-z0-9_.]+)\}\}", replace, template)


def language_switch(page, lang, root, locales):
    links = []
    for other in LANGUAGES:
        current = ' aria-current="true"' if other == lang else ""
        links.append(
            f'<a href="{root}{folder(other)}{page}" hreflang="{other}" lang="{other}" '
            f'title="{locales[other]["language_name"]}" data-lang="{other}"{current}>{other.upper()}</a>'
        )
    label = locales[lang]["nav"]["language"]
    # Remember an explicit choice, so the home page doesn't redirect again.
    script = ("<script>document.querySelectorAll('.lang-switch a').forEach(a=>a.addEventListener('click',()=>"
              "{try{localStorage.setItem('livro-lang',a.dataset.lang)}catch(e){}}))</script>")
    return f'<span class="lang-switch" role="group" aria-label="{label}">{"".join(links)}</span>{script}'


def alternates(page, root):
    base = SITE_URL or root
    lines = [f'<link rel="alternate" hreflang="{l}" href="{base}{folder(l)}{page}">' for l in LANGUAGES]
    lines.append(f'<link rel="alternate" hreflang="x-default" href="{base}{page}">')
    return "\n".join(lines)


# On the default-language home page only: a first-time visitor whose browser
# prefers another supported language is sent to it. Picking a language in the
# switcher is remembered and wins from then on.
REDIRECT = """<script>
(function () {
  try {
    if (localStorage.getItem('livro-lang')) return;
    var supported = %s;
    var prefs = navigator.languages || [navigator.language || ''];
    for (var i = 0; i < prefs.length; i++) {
      var code = (prefs[i] || '').slice(0, 2).toLowerCase();
      if (code === %s) return;
      if (supported.indexOf(code) !== -1) { location.replace(code + '/' + location.hash); return; }
    }
  } catch (e) {}
})();
</script>"""


def footer(lang, root, strings):
    f, legal = strings["footer"], strings["legal"]
    return f"""<footer class="site-footer">
  <div class="wrap footer-row">
    <span>{f["made"]} · {f["contact"]}: <a href="mailto:{CONTACT_EMAIL}">{CONTACT_EMAIL}</a></span>
    <nav aria-label="{f["label"]}">
      <a href="index.html#about">{f["about"]}</a>
      <a href="index.html#pricing">{strings["nav"]["pricing"]}</a>
      <a href="index.html#team">{strings["nav"]["team"]}</a>
      <a href="terms.html">{legal["terms"]}</a>
      <a href="privacy.html">{legal["privacy"]}</a>
    </nav>
  </div>
</footer>"""


def main():
    locales = {l: json.loads((ROOT / "locales" / f"{l}.json").read_text(encoding="utf-8")) for l in LANGUAGES}
    templates = {name: (ROOT / "templates" / name).read_text(encoding="utf-8") for name in {"index.html", "legal.html"}}

    for lang in LANGUAGES:
        strings = locales[lang]
        root = "" if lang == LANGUAGES[0] else "../"
        out_dir = ROOT / folder(lang)
        out_dir.mkdir(exist_ok=True)

        for page, template, source, title_key in PAGES:
            values = {
                "lang": lang,
                "root": root,
                "langs": language_switch(page, lang, root, locales),
                "alternates": alternates(page, root),
                "footer": footer(lang, root, strings),
                "redirect": "",
            }
            if source is None:
                if lang == LANGUAGES[0]:
                    values["redirect"] = REDIRECT % (json.dumps(LANGUAGES[1:]), json.dumps(lang))
            else:
                docx = ROOT / "legal" / lang / source
                title, blocks, toc = convert(docx)
                values.update(
                    page_title=lookup(strings, title_key),
                    source=f"{lang}/{source}",
                    title=html.escape(title),
                    toc="\n".join(
                        '      <li><a href="#{}">{}</a></li>'.format(a, html.escape(re.sub(r"^[0-9]+[.)]\s*", "", t)))
                        for a, t in toc
                    ),
                    body="\n".join("    " + b for b in blocks),
                )
            (out_dir / page).write_text(fill(templates[template], values, strings), encoding="utf-8")
            print(f"wrote {folder(lang)}{page}")


if __name__ == "__main__":
    main()
