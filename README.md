# Site Content Crawler

Crawls every page of a website and catalogs its content into an Excel spreadsheet.

For each page it captures:

| Column | Notes |
|---|---|
| Page URL | Final URL after redirects |
| Status | HTTP status code (broken pages are listed too) |
| Page Title | Text of the `<title>` tag, plus its length |
| Meta Description | `<meta name="description">` (falls back to `og:description`), plus its length |
| H1 | Main heading. If a page has more than one H1, they're joined with ` \| ` |
| Body Copy | All of the page's visible copy, one paragraph per line, with subheadings marked like `[H2] Our Services` |
| Word Count | Words in the body copy |
| Image Alt Text | Each image's alt text, one per line |
| Images / Images Missing Alt | Image counts, useful for accessibility and SEO audits |

A second sheet, **Images**, has one row per image (Page URL, Image URL, Alt Text) and shows `(missing)` for images that have no alt text.

## Setup

```bash
pip install -r requirements.txt
```

Requires Python 3.9+.

## Usage

```bash
python crawler.py https://www.example.com
```

This writes `example.com-content.xlsx`. Common options:

```bash
python crawler.py https://www.example.com -o example.xlsx   # choose output file (.xlsx or .csv)
python crawler.py https://www.example.com --max-pages 2000  # default limit is 500
python crawler.py https://www.example.com --path-prefix /blog/
python crawler.py https://www.example.com --include-subdomains
python crawler.py https://www.example.com --keep-boilerplate   # include nav/header/footer text
python crawler.py https://www.example.com --all-images         # include logos/icons in header & footer
python crawler.py https://www.example.com --delay 1            # be gentler on the server
```

Run `python crawler.py --help` for every option.

## How it works

- Starts at the URL you give it and seeds the queue from `sitemap.xml` (and any sitemaps listed in `robots.txt`). Then it follows every link that stays on the same domain. `www.` and the bare domain count as the same site.
- Obeys `robots.txt` (turn this off with `--ignore-robots`) and waits 0.5 seconds between requests by default.
- Ignores URL fragments and tracking parameters (`utm_*`, `gclid`, `fbclid`, …) so the same page isn't listed twice. Skips non-HTML files such as PDFs and images.
- **Body copy** comes from the page's `<main>` region when it has one, otherwise from `<body>`. Site-wide navigation, headers, footers, sidebars, cookie banners and scripts are removed, so each row shows only that page's own content. The H1/H2/H3 headings stay in the text, marked inline.
- Excel limits a cell to 32,767 characters. Body copy longer than that is cut off and ends with a `[truncated]` note. Use `-o file.csv` if you need the full text of very long pages.

## Limitations

- The crawler reads the HTML the server sends and does not run JavaScript. Sites that build their content in the browser (some React/Angular single-page apps) may come back with little or no body copy.
- Pages behind a login aren't reachable.

## Tests

```bash
pip install pytest
python -m pytest tests
```
