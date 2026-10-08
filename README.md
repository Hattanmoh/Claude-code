# Language Champion · Program Toolkit

A 23-slide PowerPoint for launching **Language Champion**, AMP-2's peer-to-peer language program: colleagues who speak Arabic, Turkish, French, Urdu or Spanish teach their language to fellow employees.

![Language Champion title slide](docs/img/cover.jpg)

**[Download the deck](output/Language_Champion_Program_Toolkit.pptx)** — native, fully editable PowerPoint (16:9), with speaker notes on every slide.

## What's inside

| Part | Slides | Content |
| --- | --- | --- |
| Overview | 1–2 | Title slide and a guide to the toolkit |
| 1 · The pitch | 3–13 | Why now, the five languages, peer-to-peer learning, how the program works, the Champion role and responsibilities, benefits, selection process, timeline, call to action |
| 2 · Announcement email | 14–16 | Send plan, email mock-up and its key sections; the full email text is in the speaker notes of slide 15 |
| 3 · Selection criteria | 17–22 | Committee and eligibility gates, 100-point scorecard, rating scale, 10-minute conversation, decision workflow, printable scoring sheet |
| Launch | 23 | Checklist of every placeholder to complete before sending |

![All 23 slides](docs/img/all-slides.jpg)

## Before you present

Replace the bracketed placeholders:

- `[INSERT REGISTRATION LINK]` and the QR code on slide 13
- `[INSERT LEARNER REGISTRATION LINK]`
- `[Sender name and title]` and `[HR contact name and email]`
- Committee member names on slide 18
- `[session length]` on slide 9 and the kickoff date
- The rationale for the five languages on slide 5

## Design

The deck uses the AMP-2 Lucid house style: navy `#22335A`, light sand `#F4EEE5` and orange `#D9772B`, with Georgia headings and Arial body text. The colours are built into the PowerPoint theme, so they appear in the colour picker when you edit.

The full specification — palette, type scale, grid, components and rules — is in [docs/design-spec.md](docs/design-spec.md). You can paste it into any AI tool as a design brief to produce new decks in the same style.

## Rebuild or edit the deck

The deck is generated from code, so a change to the content or design is one edit and one command.

```bash
npm install
npm run build          # writes output/Language_Champion_Program_Toolkit.pptx
```

Requires Node.js 18 or later.

| To change | Edit |
| --- | --- |
| Slide text, layout or order | `src/build.js` (one block per slide, in order) |
| Colours and fonts | the `THEME` object at the top of `src/build.js` |
| The announcement email | `content/announcement-email.txt` |
| Logo | `assets/lucid-navy.png` and `assets/lucid-white.png` |

## Repository layout

```
assets/        Lucid wordmark in navy (light slides) and white (dark slides)
content/       Announcement email text, inserted into the speaker notes
docs/          Design specification and preview images
output/        The built PowerPoint
src/           Deck generator (pptxgenjs) and theme-colour step
```
