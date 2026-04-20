# Rockville Landing Page QA Checks

- SEO basics: Updated `<title>`, meta description, canonical URL, localized H1/H2 hierarchy reviewed manually.
- Schema: Embedded JSON-LD graph for LocalBusiness, Person (Ari Baylor), and FAQPage validated for syntax via visual inspection.
- CTAs: Primary `Book Your Free Consultation` buttons use booking URL with GA4 click events; modal triggers carry GA4 events; sticky nav and mobile bottom bar present.
- Content/alt: Hero bullets, intro paragraphs, service area blurbs, trainer spotlight, and FAQs localized for Rockville; new alt text references Rockville scenes.
- Links/forms: Booking, tel, mailto, and internal anchors verified for accuracy; application form validation and redirect logic preserved.
- Console: Static code review confirms no missing selectors after JS updates; no new external dependencies introduced.
