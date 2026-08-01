# hermitaudio.github.io

Personal archive site for Terje Sandstrøm's audio-design history: Electrocompaniet
1975-79 (the "Otala" 25W amplifier and its history), amplifier schematics and
troubleshooting guides, theory articles, an M.Sc thesis on weak nonlinearities, and
misc audio-design write-ups.

## Stack

- **Hugo** static site, theme [Relearn](https://github.com/McShelby/hugo-theme-relearn)
  (docs/wiki style, sidebar tree nav) as a git submodule under `themes/hugo-theme-relearn`.
- Deployed to **GitHub Pages via GitHub Actions** (`.github/workflows/hugo.yaml`),
  same pattern as the sibling blog at `C:\repos\blog\jul26blog`. Default branch is
  `master`.
- This is a **structured reference/archive site**, not a blog — organize content by
  topic section (history, schematics, theory, thesis, audio-design), not by post date.

## `Legacy/`

The entire pre-Hugo site (old Jekyll setup + Microsoft FrontPage/ASP-era files) was
moved here verbatim when the Hugo migration started. It is the source of truth for
content that still needs to be migrated into `content/`. Treat it as read-only
reference material — don't edit it, just read from it when porting a page over.

Some files in `Legacy/` are dead weight from the FrontPage/ASP era (`default.asp`,
`global.asa`, `compctrl.htm`, `mod_pics.aspx`, `_borders/`, `_fpclass/`, `_themes/`,
`_vti_cnf/`, `_vti_pvt/`) — these don't need to be ported, just left archived.

## Content migration

Migration from `Legacy/` into `content/` is incremental, page by page — not a
bulk conversion. When porting a page:
1. Read the source `.htm`/`.md` file in `Legacy/`.
2. Write it as a proper Hugo content page (front matter, Markdown body) in the
   right section under `content/`.
3. Move referenced images alongside it (page bundle) or into `static/`, whichever
   fits how Relearn resolves the theme's image paths.
4. Leave the original in `Legacy/` untouched until the whole site is ported and
   verified — don't delete Legacy content as you go.

See `.claude/skills/audio-design-content/SKILL.md` for the content conventions,
and `.claude/skills/hugo-github-pages-setup/SKILL.md` for the Hugo+Pages toolchain
if it ever needs to be reproduced elsewhere.

## Session log

`session_log.md` at the repo root tracks what's been done and decided across
sessions. Append to it, don't rewrite history in it.
