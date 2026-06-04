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

**Recurring:** Hit again on 2026-06-02 with generated slides appearing white/blank after the copy/paste path. PowerPoint `Slide.Copy` is clipboard-based and can be unreliable when the source deck is hidden and the destination view is being controlled programmatically. Use `Slides.InsertFromFile` as the primary import path for downloaded Docmee PPTX files, print the import method/download URL for comparison, and keep copy/paste only as a fallback.

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

**Recurring:** Hit again on 2026-06-02 with screenshots still showing the old blank-card UI after code had been pushed. For installer-delivered VSTO changes, bump `OfficeAgent` `ProductVersion`/`ProductCode`/`PackageCode` and show a visible task-pane build marker so testers can confirm they installed the new package.

## 2026-06-02 Docmee Template List Needs Fallback And Diagnostics

**Problem:** After outline generation, the task pane could show `模板加载失败。` with an empty disabled template selector while the generated outline text remained visible.

**Root Cause:** `LoadTemplatesAsync` treated any template-list exception as a terminal UI failure. It disabled the selector, hid the actual exception behind a short message box, and had no built-in demo templates, so users could not continue to generate/import PPT even when outline generation had succeeded.

**Solution:** Add known Docmee demo templates as a fallback list, centralize template population so live and fallback templates render identically, show `模板接口失败，已使用内置模板 N 个。`, and append the real exception chain into the output box for diagnosis.

**Recurring:** Hit again on 2026-06-02 with users reporting the task pane could still feel stuck and template covers still did not appear. `ListTemplatesAsync` still used the default 5-minute client timeout before falling back, and template cards still used `PictureBox.LoadAsync`/remote `ImageLocation`, which is opaque and unreliable inside Office task panes. Use a short timeout for the template-list call, render selectable cards before any image request, and load covers through a controlled 5-second `HttpClient` download path.

**Recurring:** Hit again on 2026-06-02 when clicking refresh made PowerPoint appear frozen. Even after lowering timeouts, `LoadTemplatesAsync` and `BeginLoadTemplateCovers` still initiated HTTP work from the task-pane async methods, and the shared client awaits could capture the Office UI synchronization context. Guard refresh against overlapping loads, yield once so the UI repaints, start template-list and cover HTTP work from `Task.Run`, and use `ConfigureAwait(False)` inside the Docmee client methods.

## 2026-06-02 VB Local Variables Can Shadow Type Names

**Problem:** GitHub Actions failed compiling `PowerPointAi/ThemePptTaskPane.vb` with `BC30980: Type of 'image' cannot be inferred from an expression containing 'image'` and `BC42104`.

**Root Cause:** VB is case-insensitive. A local variable named `image` in `Dim image = Await Task.Run(...)` shadowed the imported `System.Drawing.Image` type used inside the same initializer (`Image.FromStream`, `CType(..., Image)`), so the compiler tried to infer the variable from an expression that referenced itself.

**Solution:** Rename the local variable to `coverImage`, explicitly type it as `System.Drawing.Image`, and fully qualify `System.Drawing.Image.FromStream`/`CType(..., System.Drawing.Image)`. Add a static verification check preventing `Dim image = Await Task.Run` from returning.

**Recurring:** Hit again on 2026-06-03 while adding Markdown preview rendering. A function parameter named `markdown` could shadow the imported `Markdig.Markdown` type because VB is case-insensitive. Use the fully qualified `Markdig.Markdown.ToHtml(...)` call when rendering Markdown, especially inside methods that also accept a `markdown` parameter.

## 2026-06-03 PowerPoint COM Must Stay On The Office UI Thread

**Problem:** While adding the single-slide replacement workflow, it was tempting to wrap the whole operation in `Task.Run` so generation would feel asynchronous.

**Root Cause:** PowerPoint VSTO automation uses Office COM objects that are tied to the Office STA/UI thread. Running slide selection, insertion, deletion, or formatting from a background thread can cause hangs, invalid COM access, or unpredictable UI state.

**Solution:** Keep the Ribbon handler as `Async Sub`, await only the model/network call, and perform all PowerPoint COM operations after the await on the captured Office UI thread. Do not put Office slide manipulation inside `Task.Run`.

## 2026-06-03 Static Verification Must Follow Configurable Integration Changes

**Problem:** After changing the Docmee client from hard-coded demo URL/token constants to configurable endpoint/token helpers, `node scripts/verify-docmee-theme-ppt.js` failed because it still expected the old `Private Shared ReadOnly UpdatePptTemplateEndpoint` field shape.

**Root Cause:** The verification script was asserting implementation details instead of the stable contract: configured base URL, configured token, and the required Docmee paths.

**Solution:** Update the script to verify the behavior-level contract: Docmee settings are exposed through `ConfigSettings` and `app.config`, endpoint properties are built from the configured base URL, requests use the shared token helper, and no request code hard-codes `DemoToken`.

## 2026-06-03 Docmee Editable Document Outlines Should Use Type 2

**Problem:** The task pane's document-to-PPT flow mapped `.doc` and `.docx` uploads to Docmee `type=4`, then called `generateContent` to produce editable Markdown.

**Root Cause:** Docmee `type=4` is for Word precise conversion, and the official V2 docs say the `prompt` parameter is ignored for task types other than 1, 2, 5, and 6. The task pane flow is not direct Word conversion; it needs an uploaded document task that can generate editable Markdown and honor the optional user prompt.

**Solution:** Use Docmee `type=2` for the task pane's document upload flow, including Word files, so `generateContent` can produce editable Markdown and keep prompt behavior consistent.

## 2026-06-03 Docmee Upload Filter Must Match Official Formats

**Problem:** The task pane's document picker only allowed Word, PDF, TXT, and Markdown files, even though Docmee V2 upload-file tasks support more document inputs.

**Root Cause:** The UI filter was hard-coded narrowly and the client did not locally validate file extensions before upload, so users could not pick supported PPT/Excel/HTML/ebook/mind-map inputs from the normal picker and unsupported files would fail later at the API boundary.

**Solution:** Expand the picker to Docmee's supported upload formats (`doc/docx/pdf/ppt/pptx/txt/md/xls/xlsx/csv/html/epub/mobi/xmind/mm`) and add a local unsupported-format error before creating the upload task.

## 2026-06-03 App Config Defaults Can Mask Environment Overrides

**Problem:** Docmee environment variables would not take effect because `PowerPointAi/app.config` contains non-empty demo defaults for the Docmee base URL and token.

**Root Cause:** `ConfigSettings.GetDocmeeApiBaseUrl` and `GetDocmeeToken` checked app settings before environment variables. Since app settings always contained `https://test.docmee.cn` and `ak_demo`, the environment fallback path was unreachable.

**Solution:** Read Docmee settings in this order: runtime value, user-saved plugin settings, environment variables, app.config, then demo defaults. Add static verification that environment variables appear before app settings in the fallback chain.

**Recurring:** Hit again on 2026-06-03 with the Docmee settings dialog. Saving a blank API base URL used `FirstNonEmpty(apiBaseUrl, DefaultDocmeeApiBaseUrl)`, which persisted the demo URL into the user settings file and could again mask environment variables. Save the user's literal trimmed value instead, and only apply defaults during read/fallback.

## 2026-06-03 PPT Fill Mode Must Not Reuse Nonblank Text Filtering

**Problem:** The PowerPoint `填充` text optimization mode could miss an empty direct text selection/cursor target and blank table cells, even though filling blank placeholders is the core expected workflow.

**Root Cause:** The selected-text and table-cell target collectors reused the normal optimization filter `Not String.IsNullOrWhiteSpace(...)`, while only the selected-shape branch respected the `allowBlankTextFrame` flag.

**Solution:** Apply `allowBlankTextFrame OrElse Not String.IsNullOrWhiteSpace(...)` consistently to direct text selections, selected shapes, and selected table cells, and add static regression checks for each collection path.

## 2026-06-03 Docmee Async Continuations Should Stay Off Office UI Context

**Problem:** Docmee generation could still feel sluggish in Office even after moving template loading and cover downloads to background tasks, because long HTTP/SSE continuations could resume on the Office UI synchronization context.

**Root Cause:** Several Docmee client awaits for `createTask`, document upload, `generateContent` SSE reading, `generatePptx`, `downloadPptx`, `updatePptTemplate`, and PPTX byte download did not use `ConfigureAwait(False)`.

**Solution:** Add `ConfigureAwait(False)` to Docmee client network and stream awaits while keeping all PowerPoint COM import/replacement/formatting operations on the plugin UI thread, and add static checks so future Docmee HTTP paths do not regress.

## 2026-06-03 Replacement Slide Must Not Fake AI Output

**Problem:** PowerPoint `替换单页` could appear to work without model configuration because the workflow used the user's requirement text as fallback slide content.

**Root Cause:** `GenerateReplacementSlideTextAsync` returned `requirement.Trim()` when the model API settings were missing or the model returned an empty response, hiding configuration failures and producing a non-generated slide.

**Solution:** Require API URL, API Key, and model name before replacement generation, throw a clear error when the model returns no usable content, and keep regression checks that forbid `Return requirement.Trim()` in the replacement generation path.

## 2026-06-04 Streaming UI Callbacks Must Handle Disposed Task Panes

**Problem:** After moving Docmee stream reading off the Office UI synchronization context, late Markdown outline chunks could still arrive after the PowerPoint task pane had been disposed.

**Root Cause:** `AppendOutlineStreamText` marshaled background stream chunks with `BeginInvoke`, but it did not first check whether the task pane or output box had already been disposed.

**Solution:** Return early when `Me.IsDisposed` or `_outputBox.IsDisposed` is true, and add a static regression check so streaming callbacks keep this guard.

**Recurring:** Hit again on 2026-06-04 with a narrower race: several background callbacks still called raw `BeginInvoke` after an `InvokeRequired` check, so the pane could be disposed between the check and the marshal call. Centralize marshaling in `BeginInvokeIfAlive`, catch disposed/invalid handle exceptions, and dispose transferred images when marshaling fails.
