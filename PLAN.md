# Plan for Rockville Landing Page

- [ ] Review `LIVE McLean Landing Page HTML 10-17.html` structure, assets, and scripts to scope edits for Rockville variant.
- [ ] Draft Rockville-localized copy (hero, intro, neighborhoods, FAQs, trainer spotlight) referencing provided neighborhoods, landmarks, and commuting context.
- [ ] Update CTA labels/links and instrumentation to use Rockville booking URL while preserving GA4 tracking.
- [ ] Implement localized content within new `rockville.html`, ensuring SEO elements, schema, accessibility, and CTA placements meet requirements.
- [ ] Add supporting docs (`CHECKS.md`) and verify sticky header, mobile CTA bar, links, and console cleanliness.

## Acceptance Criteria

- Unique Rockville-focused content delivered for hero, intro, "Where We Train," Trainer Spotlight, and FAQs.
- All CTAs point to the Rockville booking URL with GA4 events firing on clicks and form submits.
- On-page SEO: updated title/meta, headings, canonical, localized alt text, and JSON-LD (LocalBusiness, FAQPage, Person).
- Accessibility maintained (semantic headings, labeled forms, alt text) with no console errors or broken links.
- Minimal structural changes; existing CSS/JS reused without introducing new libraries.

> NOTE: `ARI_HEADSHOT_URL` not provided—will implement a clearly labeled placeholder image block in Trainer Spotlight.
