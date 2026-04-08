# Student Assessment UI Audit (Implementation-Grounded)

Date: 2026-04-08

## 1. Executive Summary

The student interface feels high-effort but low-trust because it combines exam-critical workflows with too many decorative motifs (glass, glow, gradients, emoji states, animated interstitials) and too much state-specific one-off styling. The result is an inconsistent visual language that reads more like a polished demo than a disciplined testing surface.

## 2. Files and Render Path

- `StudentApp.html`: owns almost all student CSS, DOM shell, and client-side state/render logic.
- `Code.gs`: routes student traffic via `doGet` and exposes top-level wrappers used by `google.script.run`.
- `Data.gs`: provides session/question/status payloads that drive student UI states, including snapshot-first questions, reveal details, lock/pause/timer state, and summary config.
- `Grading.gs`: influences student-visible post-submit/reveal behavior via score/feedback sync to response rows and grading status.
- `TeacherApp.html`: useful comparison; similar token family but tighter radii/shadows and denser, more utilitarian admin framing.

## 3. High-Confidence Visual Problems

1. **Competing motifs (exam UI + theatrical interstitials).**
   - The core exam shell and question cards are mixed with a full-screen “get ready” stage (joke/quote modes, ambient glow, animated progress line), plus animated “scoring” and “waiting” cards.
   - This shifts tone from assessment trust to presentation theatrics.

2. **Glass/blur overuse across too many layers.**
   - Join card, header, question card, stimulus card, modals, overlays, calculator button/window, toasts, completion card all use transparent + blur treatments.
   - Reduced visual stability hurts readability and seriousness during timed testing.

3. **State fragmentation from inline style strings.**
   - `renderQ()` builds many separate visual states with large inline style blobs.
   - Small tweaks become inconsistent across states; visual drift accumulates.

4. **Join/fullscreen/assessment/done feel like different products.**
   - Join and fullscreen prompts are cinematic; assessment shell is dashboard-like; completion card is celebratory.
   - There is no single restrained “exam mode” language.

5. **Over-animated feedback loops for high-stress context.**
   - Pulse, bounce, gradient-shift, drop-in, fade-in, ring animations are layered across multiple states.
   - Motion load can feel distracting where cognitive focus should be on the item stem and choices.

## 4. Design System Diagnosis

- **Too many surface definitions** (glass card variants, alert variants, gradient strips, glows).
- **Weak restraint in interaction styling** (hover lift/shadow on core answer options, decorative transitions for operational states).
- **Typography personality drift** (Inter + Outfit + JetBrains + decorative quote serif mark usage).
- **Monolithic single-file architecture** (3,300+ lines) makes style governance hard, especially with render-time inline styles.
- **State-first visuals vs system-first visuals**: each state invents new ornamentation instead of reusing a strict component language.

## 5. Code-Level Evidence

- Heavy visual-token ambition in root and body (large radii, layered shadows, global gradient background).
- Join card and action button use deep blur + gradient/hover glow.
- Header/timeline/question cards/stimulus/textarea/options all layer blur/shadow/hover movement.
- Transition stage includes mode-specific glows/colors/animations and joke content.
- `renderQ()` includes many inline-styled sublayouts for reveal, locked, meta, scoring, waiting, and holding states.

## 6. Recommended Design Direction

Adopt a strict **Exam Calm** profile:

- One base surface style (opaque, subtle border, minimal shadow).
- One radius scale (e.g., 10/12 only).
- Remove decorative gradients except primary action button fill.
- Reduce typography to Inter body + one display family for headings (no decorative quote styling).
- Flatten overlays: solid dim backdrop + simple panel.
- Keep only essential motion (very short opacity fade).
- Preserve all security/timer/lock behavior; simplify only presentation.

## 7. Exact Code-Level Recommendations by File

### `StudentApp.html`
- Replace glass-heavy tokens (`--s`, `--s2`, high blur, large shadow/radius) with flatter, opaque surfaces.
- Remove hover lift on `.mc-opt`; keep border-color transition only.
- Replace animated “get ready” joke overlay with brief neutral transition card (or optional teacher toggle default-off).
- Consolidate state cards (`locked`, `scoring`, `waiting`, `done`) under one reusable `.state-card` style.
- Move render-time inline style fragments into named classes; keep `renderQ()` focused on structure/state logic.
- Normalize join/fullscreen/pause/lock/done into same panel system.

### `Code.gs`
- No behavior change needed; keep route/wrapper boundaries stable.

### `Data.gs`
- No behavior change needed; snapshot/reveal/timer payload model is correct for stable student rendering.

### `Grading.gs`
- No behavior change needed; student-facing trust benefit comes from UI treatment, not grading pipeline.

### `TeacherApp.html`
- Use as design-control reference: tighter radius/shadow baseline can be ported to student shell.

## 8. Top 15 Highest-Value Fixes

1. Remove get-ready joke/quote theatrical stage from mandatory flow.
2. Flatten all primary student surfaces (header, cards, footer) to opaque.
3. Standardize one state card component for lock/pause/scoring/waiting/done.
4. Remove hover translate effects on answer options.
5. Reduce gradients to primary CTA only.
6. Drop glow accents and animated top bars in state cards.
7. Reduce radius from 20–32px range to a consistent 10–12px scale.
8. Cut toast styling to plain high-contrast notification chip.
9. Simplify timeline visuals (fewer rings/shadows/tags styles).
10. Normalize spacing scale (8/12/16/24/32).
11. Restrict heading styles; reduce display-size spikes.
12. Replace emoji-forward state headers with clear plain-language labels.
13. Consolidate inline styles into reusable CSS classes.
14. Move style-heavy render branches into helper template functions.
15. Final pass for density consistency between header, item body, and footer controls.

## 9. PR-Ready Patch Plan

1. Define restrained token set (`Exam Calm`) in `:root`.
2. Flatten page/background/surfaces and remove pervasive blur.
3. Normalize typography and spacing scale.
4. Refactor question/stimulus/answer card styles to one system.
5. Introduce unified state panel for pause/lock/wait/score/done.
6. Reduce decorative motion to minimal fades only.
7. Replace inline render styles with semantic classes.
8. Run consistency pass across join → exam → submit → completion.

## 10. Draft Patch Suggestions

- Add utility classes: `.panel`, `.panel--state`, `.panel--warning`, `.panel--danger`, `.muted`, `.stack-*`.
- In `renderQ()`, swap:
  - `'<div class="q-card" style="text-align:center; padding:64px 48px; ...">'`
  - to `'<section class="panel panel--state state-wait">'` style blocks.
- Replace `.mc-opt:hover { transform: translateY(-2px); box-shadow: ... }` with border-color-only hover.
- Change `.q-card`, `.stim-card`, `.a-hdr`, `.a-foot`, `.modal-card`, `.done-card` to shared border/shadow/radius primitives.
- Gate `triggerGetReadyAnimation` behind a config flag and default to a short neutral transition.
