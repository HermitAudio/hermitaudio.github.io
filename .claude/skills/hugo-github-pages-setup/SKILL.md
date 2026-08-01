---
name: hugo-github-pages-setup
description: Stand up a new Hugo static site (or scaffold one into an existing repo) deployed to GitHub Pages via GitHub Actions. Use when asked to set up Hugo, migrate a site to Hugo, or wire up Pages deployment for a Hugo repo.
---

# Hugo + GitHub Pages setup

Checklist for standing up a Hugo site with an Actions-based Pages deploy, distilled
from `hermitaudio.github.io` and its sibling `jul26blog`.

## 1. Site init

```bash
hugo new site . --force   # --force if scaffolding into a non-empty repo
```

If old site content already lives at the repo root, move it out of the way first
(e.g. into a `Legacy/` folder via `git mv`) before running `hugo new site .` — Hugo
won't overwrite existing files, but a cluttered root makes the new structure
confusing.

## 2. Theme as a git submodule

```bash
git submodule add https://github.com/<owner>/<theme-repo>.git themes/<theme-name>
```

Set `theme = "<theme-name>"` in `hugo.toml`. Using a submodule (not a vendored copy)
keeps the theme updatable and matches how CI checks it out (`submodules: recursive`
in the checkout step below).

Theme choice depends on the site's shape:
- **Blog / post feed** → Ananke (`theNewDynamic/gohugo-theme-ananke`) — simple,
  taxonomy + menu driven, used by `jul26blog`.
- **Structured reference/docs/wiki site** (persistent sidebar nav, sections rather
  than a chronological feed) → Relearn (`McShelby/hugo-theme-relearn`) — sidebar
  tree, built-in search, used by `hermitaudio.github.io`.

## 3. `hugo.toml` essentials

```toml
baseURL = 'https://<user-or-org>.github.io/'   # or custom domain if CNAME is set
languageCode = 'en-us'
title = '<Site Title>'
theme = '<theme-name>'

[markup]
  [markup.goldmark]
    [markup.goldmark.renderer]
      unsafe = true   # allow raw HTML in Markdown, useful when porting legacy .htm
```

Add menu entries / section config per the chosen theme's docs — Ananke uses
`[[menu.main]]` blocks; Relearn drives most of its nav from `content/` structure
plus per-page `weight` front matter.

## 4. `.gitignore`

Hugo build output and caches should never be committed:

```
/public
/resources
.hugo_build.lock
```

## 5. GitHub Actions workflow

`.github/workflows/hugo.yaml` — install Hugo extended, checkout with
`submodules: recursive`, build with `--gc --minify`, deploy via the official
Pages actions:

```yaml
name: Deploy Hugo site to Pages

on:
  push:
    branches: [master]   # match the repo's actual default branch
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: "pages"
  cancel-in-progress: false

jobs:
  build:
    runs-on: ubuntu-latest
    env:
      HUGO_VERSION: 0.164.0
    steps:
      - name: Install Hugo CLI
        run: |
          wget -O ${{ runner.temp }}/hugo.deb https://github.com/gohugoio/hugo/releases/download/v${HUGO_VERSION}/hugo_extended_${HUGO_VERSION}_linux-amd64.deb \
          && sudo dpkg -i ${{ runner.temp }}/hugo.deb
      - name: Checkout
        uses: actions/checkout@v7
        with:
          submodules: recursive
          fetch-depth: 0
      - name: Setup Pages
        id: pages
        uses: actions/configure-pages@v6
      - name: Build with Hugo
        env:
          HUGO_ENVIRONMENT: production
        run: hugo --gc --minify --baseURL "${{ steps.pages.outputs.base_url }}/"
      - name: Upload artifact
        uses: actions/upload-pages-artifact@v5
        with:
          path: ./public

  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    needs: build
    steps:
      - name: Deploy to GitHub Pages
        id: deployment
        uses: actions/deploy-pages@v5
```

Pin action versions to a commit SHA (with the version as a trailing comment) if the
repo's other workflows do — check for a `renovate.json` / Dependabot config that
expects it.

## 6. GitHub repo settings (manual — cannot be done via git)

Repo → **Settings → Pages → Build and deployment → Source**: set to
**"GitHub Actions"**. If the repo previously deployed Jekyll from a branch, this
setting needs to be switched explicitly — the old branch-deploy config doesn't
disable itself. Always call this out to the user rather than assuming it's set.

If the site uses a custom domain, a `CNAME` file at the repo root (containing just
the domain) is still required even with Actions-based deploy — Hugo's `baseURL`
alone doesn't configure Pages' domain routing.

## 7. Local verification

```bash
hugo server -D        # draft content included, check nav/build live
hugo --gc --minify    # production-mode build check, same flags as CI
```

Confirm `public/` is populated and there are no broken theme/shortcode references
before committing.
