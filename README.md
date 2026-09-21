# Melbits Astro Website Template

A modern, production-ready Astro website for **melbits.com.au** — rebuilt with improved content, design system, and performance.

## Brand

- Primary Color: `#0AB0E8`
- Fonts: **Sora** (headings/display) + **DM Sans** (body)
- Dark background: `#0b1220`

## Tech Stack

- [Astro 4](https://astro.build) — static site generator
- Vanilla CSS with CSS custom properties (no external CSS framework)
- Google Fonts (loaded via `@import`)
- No JavaScript frameworks (plain JS for nav interactions)

## Project Structure

```
src/
  layouts/
    BaseLayout.astro       # HTML shell, imports Nav + Footer
  pages/
    index.astro            # Homepage
    contact.astro          # Contact page with form
    services/
      managed-it.astro     # Managed IT Services page (template for others)
  components/
    Nav.astro              # Sticky nav with dropdowns + mobile menu
    Hero.astro             # Homepage hero with animated cards
    WhyUs.astro            # 6-card "Why choose us" section
    Services.astro         # Services grid
    Industries.astro       # Dark-themed industries section
    Testimonials.astro     # Client testimonials grid
    CTA.astro              # Full-width CTA banner
    Footer.astro           # Footer with partners bar
  styles/
    global.css             # Design tokens, utility classes, typography
```

## Getting Started

```bash
# Install dependencies
npm install

# Start dev server
npm run dev

# Build for production
npm run build

# Preview production build
npm run preview
```

## Pages to Create

Based on current site structure, these pages still need to be built using the template pattern in `services/managed-it.astro`:

**Services**
- `/services/cloud` — Cloud Services & Azure
- `/services/microsoft-365` — Microsoft 365
- `/services/consulting` — IT Consulting
- `/services/remote-support` — Remote IT Support
- `/services/voip` — VoIP & 3CX

**Cybersecurity**
- `/cybersecurity` — Overview
- `/cybersecurity/essential-eight`
- `/cybersecurity/assessment`
- `/cybersecurity/saas-security`
- `/cybersecurity/playbooks`

**Industries**
- `/industries/accounting`
- `/industries/law`
- `/industries/medical`
- `/industries/real-estate`
- `/industries/pharmacy`
- `/industries/conveyancing`

**Company**
- `/about`
- `/blog`
- `/case-studies`
- `/faq`

## Design System

All design tokens are in `src/styles/global.css` under `:root`. Key variables:

| Variable | Value | Use |
|---|---|---|
| `--brand` | `#0AB0E8` | Primary accent |
| `--brand-dark` | `#0890c0` | Hover states |
| `--brand-light` | `#e8f8fd` | Light backgrounds |
| `--dark` | `#0b1220` | Dark sections |
| `--radius-lg` | `22px` | Card border radius |

## Deploying

Astro builds to static HTML/CSS/JS by default — deploy to:
- **Netlify**: `npm run build` → publish `dist/` directory
- **Vercel**: connect repo, auto-detects Astro
- **Cloudflare Pages**: connect repo, build command `npm run build`, output `dist`

For dynamic features (contact form), integrate with:
- Netlify Forms (add `netlify` attribute to `<form>`)
- Formspree
- Or a serverless function endpoint
