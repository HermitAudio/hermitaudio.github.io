---
name: audio-design-content
description: Conventions for writing and organizing audio-electronics content pages (history, schematics, theory, troubleshooting) on hermitaudio.github.io. Use when porting a page from Legacy/ into content/, or writing new audio-design content.
---

# Audio-design content conventions

This site documents real analog audio electronics: amplifier design history
(Electrocompaniet, the "Otala" 25W amp), schematics, troubleshooting guides,
theory (open/closed-loop frequency response, weak nonlinearities, single-stage
gain theory), and an M.Sc thesis. Accuracy and provenance matter more than prose
polish — this is reference material engineers and hobbyists will actually use to
diagnose real hardware.

## Site sections (`content/` top level)

- `history/` — the Electrocompaniet/Otala story, people involved, press coverage,
  brochures. Narrative, chronological within the section.
- `schematics/` — circuit schematics and troubleshooting guides. Each schematic
  page should state which physical unit/revision it applies to (Electrocompaniet
  amps went through hardware revisions — don't let a schematic be ambiguous about
  which one it's for).
- `theory/` — general audio theory articles, not tied to one product.
- `thesis/` — the M.Sc thesis on weak nonlinearities (scanned PDF + any
  transcribed sections).
- `audio-design/` — misc designs not tied to the Electrocompaniet history (school
  amplifier project, cooling fin calculations, AES paper on AB distortion, etc.).
- `about/` — author bio, contact.

## Front matter shape

```yaml
---
title: "<Descriptive title>"
description: "<One-sentence summary for nav/search/meta>"
weight: <int>          # controls sidebar ordering within a section (Relearn)
---
```

Add `date:` only when there's a real original-publication date worth preserving
(e.g. an AES paper year, a press clipping date) — most of this content isn't
time-sensitive and doesn't need blog-style dating.

## Porting a page from `Legacy/`

1. Read the original `.htm`/`.md` in `Legacy/` in full before touching anything —
   old FrontPage HTML is often malformed (unclosed tags, `<font>`/`<blockquote>`
   layout hacks); don't let a naive HTML→Markdown pass silently drop content.
2. Convert to clean Markdown. Keep the original wording — this is historical
   record, not a copywriting exercise. Fix only obvious typos.
3. Preserve every image, schematic, and scanned document reference. If an image
   file in `Legacy/` is referenced, copy it alongside the new page (Hugo page
   bundle: `content/<section>/<page>/index.md` + images in the same folder) rather
   than leaving it pointing back at `Legacy/`.
4. If a page duplicates another (this site has some — e.g. `abdist.md` at root and
   `wiki/ABDist.md`), reconcile into one and note in the commit message which was
   dropped and why.
5. Don't port dead infrastructure pages (old ASP scripts, ActiveX control pages,
   FrontPage-generated nav) — those aren't content, they're artifacts of the old
   hosting platform.

## Norwegian-language source material

Some original documents (school reports, press clippings, internal notes) are
in Norwegian. Default: a faithful Norwegian transcription. **An English
translation is a separate, opt-in follow-up — don't do it automatically.**

Reason this changed: the first attempt (the "Skoleforsterkeren" 31-page
school report, dense with formulas) burned a large amount of effort on the
Norwegian transcription alone; committing to translating it too, unasked, is
the kind of scope creep to avoid. Transcribe in Norwegian, ship that, and
only translate a specific page if/when asked for that page by name.

- Norwegian page: transcribe faithfully — don't silently "improve" or modernize
  the Norwegian, and don't translate technical terms unless you're certain of
  the equivalent (when unsure, leave the Norwegian term and gloss it once).
- If an English translation is requested later: full translation, not a
  summary; link the two pages to each other near the top ("Norwegian
  original" / "English translation"); formulas, tables, and figures are
  identical on both pages (only the prose is translated); as plain sibling
  content pages, not Hugo's i18n multilingual mode (this is a fixed, closed
  set of old documents, not a growing corpus — full multilingual site
  restructuring isn't warranted).
- For long/dense transcriptions with many scanned figures: render each PDF
  page as a JPEG (quality ~80-85, see `pdf-page-extraction` skill), not PNG —
  a 42-page technical scan was 45MB as PNG vs 9.6MB as JPEG, fully legible
  either way. Embed the actual page image inline next to its transcribed
  text/formulas so hand-drawn diagrams aren't lost, and use separate image
  files for figures/tables that are physically standalone pages in the
  source PDF.

## Images and scans

- Schematics and photos: keep at full legible resolution — these are often the
  *only* surviving copy of a schematic. Don't aggressively compress.
- Scanned documents (thesis, AES paper, press clippings): PDF where the source is
  a multi-page document; JPG/PNG for single schematic sheets or photos.
- Caption scanned historical documents with their real-world source (publication
  name, date, page) when known, so provenance isn't lost.
