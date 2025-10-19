# Plan for Rockville Landing Page Update

## Tasks
- [x] Review existing `mclean.html` structure and assets to mirror layout for Rockville page.
- [x] Draft localized Rockville copy (hero, intro, neighborhoods, trainer spotlight, FAQs) emphasizing conversion.
- [x] Implement `rockville.html` with localized content, SEO metadata, CTAs, and GA4 instrumentation.
- [x] Add schema JSON-LD for LocalBusiness, FAQPage, and Person entities.
- [x] Ensure accessibility, CTA placements (sticky header & mobile bar), and reuse of existing assets.
- [x] Create `CHECKS.md` documenting verification of SEO, schema, tracking, alt text, links, and console state.

## Acceptance Criteria
- Rockville content is unique, location-specific, and conversion focused across required sections.
- Primary CTA "Book Your Free Consultation" is prominent, sticky/persistent, and tracked via GA4 events.
- Page includes correct SEO metadata, canonical tag, and all required JSON-LD schemas without errors.
- Accessibility basics (headings, labels, alt text) and link integrity confirmed; no console errors expected.
- Changes remain scoped to necessary files while reusing existing CSS/JS from `mclean.html`.

> NOTE: `ARI_HEADSHOT_URL` not provided; will implement a clearly labeled placeholder image block in Trainer Spotlight.
