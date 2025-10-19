# Plan for Rockville Landing Page Updates

## Tasks
- [ ] Review existing `LIVE McLean Landing Page HTML 10-17.html` structure, assets, and scripts to identify reusable components.
- [ ] Draft localized Rockville content (hero, intro, neighborhoods, trainer spotlight, FAQs) aligned with provided guidance and Fit2Go brand voice.
- [ ] Build `rockville.html` reusing existing layout/CSS/JS while implementing localized copy, CTAs, sticky header and mobile bar, and GA4 event wiring.
- [ ] Embed required SEO metadata, canonical link, structured data (LocalBusiness, FAQPage, Person), and accessibility updates.
- [ ] Add CRO elements (primary/secondary CTAs, instrumentation) and ensure responsive behavior mirrors source page.
- [ ] Validate page for SEO basics, schema presence, CTA event tracking, alt text, link integrity, and absence of console errors; document in `CHECKS.md`.

## Acceptance Criteria
- Unique Rockville-focused content is present in hero, introduction, "Where We Train," trainer spotlight, and FAQs.
- All CTAs direct to `BOOKING_URL_ROCKVILLE` with GA4 click events; sticky header and mobile bottom bar are implemented.
- Meta tags, canonical link, headings, and JSON-LD (LocalBusiness, FAQPage, Person) meet on-page SEO requirements.
- Accessibility considerations addressed (headings, labels, alt attributes), with no broken links or console errors.
- `CHECKS.md` enumerates completed verifications.

> **NOTE:** `ARI_HEADSHOT_URL` not provided; include a clearly labeled placeholder image block in the Trainer Spotlight section.
