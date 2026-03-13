/**
 * School Terms Data
 * ─────────────────
 * Edit this file to keep term dates up to date.
 * Place it at: ~/data/schoolTerms.js in your Nuxt project.
 *
 * Each entry:
 *   label  — display name shown in the tooltip on the calendar
 *   start  — first day of the period (YYYY-MM-DD)
 *   end    — last day of the period  (YYYY-MM-DD)
 *   type   — 'term' (green band) or 'break' (amber band)
 */

export default [
  // ── 2025/26 Academic Year ──────────────────────────────────────
  { label: 'Autumn Term',   start: '2025-09-03', end: '2025-10-24', type: 'term'  },
  { label: 'Half Term',     start: '2025-10-27', end: '2025-10-31', type: 'break' },
  { label: 'Autumn Term',   start: '2025-11-03', end: '2025-12-19', type: 'term'  },
  { label: 'Christmas',     start: '2025-12-22', end: '2026-01-02', type: 'break' },

  { label: 'Spring Term',   start: '2026-01-05', end: '2026-02-13', type: 'term'  },
  { label: 'Half Term',     start: '2026-02-16', end: '2026-02-20', type: 'break' },
  { label: 'Spring Term',   start: '2026-02-23', end: '2026-04-01', type: 'term'  },
  { label: 'Easter',        start: '2026-04-02', end: '2026-04-17', type: 'break' },

  { label: 'Summer Term',   start: '2026-04-20', end: '2026-05-22', type: 'term'  },
  { label: 'Half Term',     start: '2026-05-25', end: '2026-05-29', type: 'break' },
  { label: 'Summer Term',   start: '2026-06-01', end: '2026-07-17', type: 'term'  },
  { label: 'Summer Hols',   start: '2026-07-20', end: '2026-09-04', type: 'break' },

  // ── 2026/27 Academic Year ──────────────────────────────────────
  { label: 'Autumn Term',   start: '2026-09-07', end: '2026-10-23', type: 'term'  },
  { label: 'Half Term',     start: '2026-10-26', end: '2026-10-30', type: 'break' },
  { label: 'Autumn Term',   start: '2026-11-02', end: '2026-12-18', type: 'term'  },
  { label: 'Christmas',     start: '2026-12-21', end: '2027-01-01', type: 'break' },
]