# Livro website

The Livro website in English, Turkish and German. Plain static files, hosted
on GitHub Pages straight from this folder.

| Language | Home | Terms of Service | Privacy Policy |
|---|---|---|---|
| English | `/` | `/terms.html` | `/privacy.html` |
| Türkçe | `/tr/` | `/tr/terms.html` | `/tr/privacy.html` |
| Deutsch | `/de/` | `/de/terms.html` | `/de/privacy.html` |

Every page has an EN · TR · DE switcher. On a first visit, the English home
page sends Turkish- and German-language browsers to their version; once
someone picks a language it is remembered and never redirected again.

## How it's built

The pages above are **generated** — don't edit them by hand. The sources are:

| Source | What it holds |
|---|---|
| `templates/index.html` | The home page's layout |
| `templates/legal.html` | The legal pages' layout |
| `locales/en.json`, `tr.json`, `de.json` | Every piece of home-page text, per language |
| `legal/<lang>/TermsOfService.docx` | Terms of Service, per language |
| `legal/<lang>/PrivacyPolicy.docx` | Privacy Policy, per language |
| `styles.css`, `assets/` | Shared by every page |

After changing any source:

```bash
python3 build.py
```

Then commit and push both the sources and the generated pages.

## Legal documents

This site is the home of Livro's legal documents; the app will link to the
published pages above. Each language is its own Word file, so **a change to
one document needs the same change in all three languages** — then run
`build.py`.

The Turkish and German documents are translations of the English drafts and
have the same sections in the same order; section links such as
`terms.html#section-11` point to the same section in every language. They still
need a lawyer's review.


## Adding a language

1. Copy `locales/en.json` to `locales/<code>.json` and translate it.
2. Add translated Word files under `legal/<code>/`.
3. Add the code to `LANGUAGES` in `build.py`, and run it.


## Preview locally

```bash
python3 -m http.server 8765
```
