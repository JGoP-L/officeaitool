# Pitfalls

Lessons learned from errors encountered in this project. Updated automatically by the coding agent.

## 2026-05-29 Docmee V2 Outline Shape Variants

**Problem:** The initial PowerPoint task-pane parser handled only `children` outline trees, but the Docmee V2 test environment returned a completed `generateContent` result with `overall_theme` and `pages`.

**Root Cause:** The official flow documents JSON outline generation, but live test responses can use different JSON envelope shapes for the same generated outline content.

**Solution:** Support both `children`/`items` outline trees and Docmee `pages` responses, and use `overall_theme` as a title candidate when inserting generated content into PowerPoint.

## 2026-05-29 Docmee V2 Flow Must Continue After Outline

**Problem:** The first implementation stopped after generating an outline and locally inserted simple slides, so it did not produce the actual templated Docmee PPTX.

**Root Cause:** The V2 flow was implemented only through `createTask` and `generateContent`; it missed template selection, `generatePptx`, `downloadPptx`, and importing the downloaded file into the active presentation.

**Solution:** Add template loading from `/api/ppt/templates`, generate the PPTX through `/api/ppt/v2/generatePptx`, request a downloadable file URL from `/api/ppt/downloadPptx`, download the PPTX, and insert it into the current PowerPoint with `Slides.InsertFromFile`.

## 2026-05-29 Docmee Outline Must Stream Into The UI

**Problem:** The task pane showed a blank outline area while Docmee was generating because the client waited for the full SSE response before updating the UI.

**Root Cause:** `GenerateContentAsync` used full-response buffering instead of reading the `text/event-stream` line by line. Docmee sends incremental `data:` events with `status=3` and `text` chunks before the final `status=4` outline.

**Solution:** Read the response with `ReadAsStreamAsync` and `ReadLineAsync`, invoke a progress callback for each `text` chunk, and marshal those chunks into the task pane text box with `BeginInvoke`.

## 2026-05-29 Docmee Imported Slides Need Readability Normalization

**Problem:** After importing the generated Docmee PPTX into the current PowerPoint, some body text appeared very light on pale slide backgrounds and was hard to read.

**Root Cause:** `Slides.InsertFromFile` appends slides from the generated file and returns only the inserted count; it does not provide a source-formatting or contrast option. Generated templates can contain low-contrast text after insertion into the active deck.

**Solution:** Track the newly inserted slide range from the `InsertFromFile` return value and normalize only those slides by darkening low-contrast text on light backgrounds, leaving the user's original slides untouched.
