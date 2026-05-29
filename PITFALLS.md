# Pitfalls

Lessons learned from errors encountered in this project. Updated automatically by the coding agent.

## 2026-05-29 Docmee V2 Outline Shape Variants

**Problem:** The initial PowerPoint task-pane parser handled only `children` outline trees, but the Docmee V2 test environment returned a completed `generateContent` result with `overall_theme` and `pages`.

**Root Cause:** The official flow documents JSON outline generation, but live test responses can use different JSON envelope shapes for the same generated outline content.

**Solution:** Support both `children`/`items` outline trees and Docmee `pages` responses, and use `overall_theme` as a title candidate when inserting generated content into PowerPoint.
