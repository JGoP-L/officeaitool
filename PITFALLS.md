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

## 2026-05-29 Docmee Task Pane Should Stream Markdown, Not JSON

**Problem:** The PowerPoint task pane displayed raw JSON fragments while generating the outline, but the user-facing outline area should show Markdown.

**Root Cause:** The task pane requested `outlineType=JSON` from Docmee and appended each streamed `text` chunk directly to the UI. For JSON outlines, those chunks are JSON fragments, not readable Markdown.

**Solution:** Use Docmee `outlineType=MD` for the task-pane generation flow, stream those Markdown chunks directly into the output box, and pass the same completed Markdown to `generatePptx`.

## 2026-05-29 Docmee Imported Slides Should Be Copied, Not Inserted From File

**Problem:** Users reported missing visual elements after the generated Docmee PPTX was imported into the current presentation.

**Root Cause:** `Slides.InsertFromFile` imports slides directly from a file but gives no control over source formatting preservation. Template-heavy generated slides can lose or flatten visual details when merged into an existing deck this way.

**Solution:** Open the downloaded PPTX hidden and read-only, copy each generated slide, paste it into the active presentation, then apply readability normalization only to the pasted slide objects.

**Recurring:** Hit again on 2026-05-29 with master/theme background images missing after ordinary `Slides.Paste`. The fix needs PowerPoint's keep-source-formatting paste command (`PasteSourceFormatting`) before falling back to ordinary paste.

## 2026-05-29 Docmee Markdown Stream Final Event May Contain JSON

**Problem:** The task pane failed with `Docmee 返回内容中没有可用的 Markdown result` after requesting a Markdown outline.

**Root Cause:** The live Docmee stream sends Markdown through incremental `text` chunks, but the final `status=4` event can contain a JSON outline object in `result` instead of a Markdown string. Treating that final event as required Markdown discards the already streamed Markdown and raises a false error.

**Solution:** Accumulate Markdown from all streamed `text` chunks and only use the final event when it actually contains a Markdown string; if the final event contains JSON, keep the accumulated Markdown.

## 2026-05-29 Docmee Import Must Report Zero Imported Slides

**Problem:** After selecting a template and clicking generate/import, the user could see that the operation "didn't work" without enough task-pane detail to know whether the failure was generation, download, or PowerPoint import.

**Root Cause:** `ImportPptxIntoPresentation` did not return the number of pasted slides, and the caller could report success even if no generated slides were imported.

**Solution:** Return the imported slide count, print the selected template ID, PPT ID, download URL, local path, and imported slide count in the task pane, and throw a clear error if the downloaded PPTX imports zero slides.

## 2026-05-29 Docmee V2 Tasks Lock The First PPT Template

**Problem:** Selecting a different template and clicking generate/import could still produce a PPT using the previous/default template.

**Root Cause:** The same Docmee V2 task ID was reused for final PPT generation. Live API testing showed that after one task generated PPT with template A, a second `generatePptx` call on the same task with template B returned `pptInfo.templateId` as template A and downloaded the original template file.

**Solution:** For each final generate/import action, create a fresh `type=7` Markdown task from the current outline, call `generatePptx` with the selected template ID, verify Docmee's returned `pptInfo.templateId` matches the selected template, and download with `refresh=true`.

## 2026-05-29 Docmee Import Can Leave PowerPoint On Generated Template Colors

**Problem:** Customers reported that after importing generated template slides, other PowerPoint color-related features appeared to use the generated template colors.

**Root Cause:** The import flow navigated to the end of the destination deck and used PowerPoint's keep-source-formatting paste command. After import, PowerPoint could remain on a newly pasted generated-template slide, so native color palettes and new-slide defaults were based on that slide's theme.

**Solution:** Capture the user's active slide before importing, then restore and reselect that original slide after the generated slides are pasted.

## 2026-06-02 Docmee Template Gallery Must Not Reuse The Outline Log Box

**Problem:** After outline generation and even after PPT import, the task pane could still show a large text/log box instead of the OfficePLUS-like template card gallery the user expected.

**Root Cause:** The first gallery implementation only hid the outline box once after template loading, while later diagnostics still targeted the same large output box and the import flow did not explicitly restore gallery mode.

**Solution:** Centralize task-pane view switching with `ShowOutlineOutput` and `ShowTemplateGallery`, keep the gallery visible during and after generate/import, and prevent diagnostics from bringing the large outline/log box back while gallery mode is active.

## 2026-06-02 Docmee Template Covers May Return 403

**Problem:** The task pane loaded Docmee template choices, but the template gallery appeared as blank placeholder cards and users could not tell whether a card was selectable.

**Root Cause:** The template list API returned real template metadata, but direct access to `coverUrl` resources such as `test.chatmee.cn/api/common/oss/meta-doc/ppt_template/*.png` returned HTTP 403. The UI relied too heavily on the image preview and only indicated selection with a subtle card background change.

**Solution:** Keep loading `coverUrl` when available, but render a metadata fallback preview with template name, category, style, and ID when the image cannot load. Add an explicit `选择模板` / `已选择` label so selection is visible even when remote covers are blocked.

**Recurring:** Hit again on 2026-06-02 with the live template cover path requiring the `token=ak_demo` query parameter; `ak_admin` returned JSON authentication failure instead of image bytes. Build cover image URLs through a helper that appends the demo token before assigning `PictureBox.ImageLocation`, while keeping the metadata fallback for authentication failures.
