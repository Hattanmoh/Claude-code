# Design Specification · AMP-2 Lucid House Style

The visual system used by the Language Champion Program Toolkit. Use it as a reference when editing the deck, or paste the [design brief](#reusable-design-brief) at the end into an AI tool to create new decks in the same style.

## Canvas

- 16:9 widescreen, 13.333 × 7.5 in
- Native, editable PowerPoint: real text boxes, shapes and tables; theme colours and fonts; slide layouts; named sections; speaker notes on every slide

## Colour palette

| Role | Hex | Theme slot | Used for |
| --- | --- | --- | --- |
| Navy | `22335A` | Dark 2 | Dark slide backgrounds, slide titles, emphasis cards, table headers |
| Ink | `1C2541` | Dark 1 | Body text |
| Orange | `D9772B` | Accent 1 | Icon circles, hero numbers, step markers, quotes |
| Burnt orange | `A4511A` | Accent 4 | Eyebrows, labels, buttons, small accent text |
| Light sand | `F4EEE5` | Light 2 | Alternate slide background; cards on white slides |
| Sand | `E5D7C1` | Accent 2 | Text on navy, pills, table borders |
| Slate | `4A5F86` | Accent 3 | Secondary fills, such as the weighting bar |
| Muted | `5B6478` | Accent 5 | Footer, captions |
| Taupe | `C9B79C` | Accent 6 | Connector lines, chevrons, progress bars |
| White | `FFFFFF` | Light 1 | Main content background |

Hyperlinks are burnt orange; followed hyperlinks are slate. White and light sand dominate, navy carries the structure, and orange is the sharp accent.

## Typography

Theme heading font **Georgia**, body font **Arial**. Both ship with Windows and macOS, so no fonts need embedding.

| Element | Font | Size | Colour |
| --- | --- | --- | --- |
| Title-slide headline | Georgia | 56 pt | White |
| Title-slide tagline | Georgia italic | 26 pt | Orange |
| Section-divider number | Georgia | 96 pt | Orange |
| Section-divider title | Georgia | 44 pt | White |
| Slide title | Georgia, left-aligned | 30 pt | Navy (white on dark slides) |
| Hero numbers | Georgia | 58–100 pt | Orange |
| Card headings | Georgia | 16–24 pt | Navy |
| Body text | Arial | 12–20 pt | Ink |
| Eyebrows and labels | Arial bold, uppercase, letter-spaced | 11 pt | Burnt orange (sand on dark) |
| Footer and slide number | Arial, letter-spaced | 9 pt | Muted (sand on dark) |

Arabic and Urdu text is set in Arial bold, right-to-left.

## Grid and spacing

- Side margins 0.6 in on content slides (content width 12.13 in); 0.8 in on title and divider slides
- Eyebrow at 0.42 in from the top, slide title at 0.74 in, content from 1.75 in
- Footer at 7.0 in: running text on the left, slide number on the right
- Lucid logo top right, 1.4 in wide: navy on light slides, white on dark slides
- Gutters of 0.25–0.35 in between cards

## Components

- **Cards:** rounded rectangles, 0.1 in corner radius, no outline. Light sand on white slides; white with a soft shadow on sand slides; navy for emphasis.
- **Shadow:** outer, `1C2541` at 12% opacity, 10 pt blur, 3 pt offset, straight down.
- **Signature motif:** white Font Awesome icons inside solid orange or navy circles, 0.55–1.1 in across.
- **Numbered steps:** solid circles with numerals of 14 pt or more, joined by thin taupe lines or chevrons.
- **Pills:** small rounded rectangles in light sand or burnt orange with bold 11–12 pt text.
- **Tables:** navy header row with white bold text, alternating white and light-sand rows, 0.75 pt sand borders.
- **Quotes:** Georgia italic 26–28 pt, orange, on navy.

## Slide layouts

| Layout | Background | Used for |
| --- | --- | --- |
| Title | Navy | Opening slide |
| Section divider | Navy | Start of each part |
| Content | White | Most content slides |
| Content, sand | Light sand | Alternating content slides |
| Closing | Navy | Call to action and final checklist |

## Contrast

| Pairing | Ratio |
| --- | --- |
| Ink on white / on light sand | 15.1 / 13.1 : 1 |
| White on navy | 12.4 : 1 |
| Sand on navy | 8.8 : 1 |
| Burnt orange on white / on light sand | 5.6 / 4.8 : 1 |
| Muted on white | 5.9 : 1 |
| Orange on navy | 3.9 : 1 (large text only, 26 pt and above) |
| White on orange | 3.2 : 1 (icons and bold numerals of 14 pt and above only) |

## Rules

- No accent lines under titles, and no decorative colour bars or edge stripes
- No gradients, no national flags, no stock clip-art
- Left-align body text; centre only short labels inside shapes
- Every slide has a visual element: icon, shape, table or big number
- Vary the composition from slide to slide; never plain title-and-bullets
- Nothing overflows its box: split the slide rather than shrink text below 12 pt

## Reusable design brief

```
DESIGN BRIEF: AMP-2 LUCID HOUSE STYLE (PowerPoint)

Build a native, fully editable PowerPoint (.pptx): real text boxes, shapes and tables,
not images of slides. Set the palette and fonts in the PowerPoint theme, and build the
deck on slide layouts with named sections. Speaker notes on every slide.

CANVAS: 16:9 widescreen, 13.333 x 7.5 in.

PALETTE (theme colours)
Navy 22335A (Dark 2) · Ink 1C2541 (Dark 1) · Orange D9772B (Accent 1)
Burnt orange A4511A (Accent 4) · Light sand F4EEE5 (Light 2) · Sand E5D7C1 (Accent 2)
Slate 4A5F86 (Accent 3) · Muted 5B6478 (Accent 5) · Taupe C9B79C (Accent 6) · White FFFFFF
White and light sand dominate, navy carries the structure, orange is the accent.

TYPOGRAPHY: Georgia headings, Arial body.
Title slide 56 pt · divider number 96 pt · divider title 44 pt · slide title 30 pt navy,
left-aligned · hero numbers 58-100 pt orange · card headings 16-24 pt · body 12-20 pt ink ·
eyebrows 11 pt bold uppercase burnt orange · footer 9 pt muted.

GRID: 0.6 in side margins; eyebrow at 0.42 in, title at 0.74 in, content from 1.75 in;
footer at 7.0 in with running text left and slide number right; logo top right, 1.4 in wide.
Gutters 0.25-0.35 in.

COMPONENTS: rounded cards (0.1 in radius, no outline; light sand on white slides, white with
soft shadow on sand slides, navy for emphasis). Signature motif: white icons in solid orange
or navy circles. Numbered steps in solid circles. Navy-header tables with alternating
white/light-sand rows. Orange Georgia italic quotes on navy.

LAYOUTS: navy title, navy section dividers, white and light-sand content slides
alternating, navy closing slide.

RULES: no accent lines under titles, no colour bars or stripes, no gradients, no flags.
Left-align body text. Every slide has a visual element; vary the composition.
Body text never below 12 pt; split slides rather than overflow.
```
