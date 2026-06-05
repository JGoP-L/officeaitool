# POWERPOINT AI ADD-IN

## OVERVIEW
PowerPoint-specific add-in providing Docmee-powered PPT generation, current-slide AI creation, and slide replacement capabilities. Integrates with PowerPoint through VSTO and uses shared components from ShareRibbon.

## STRUCTURE
```
PowerPointAi/
├── Ribbon1.vb           # PowerPoint ribbon implementation
├── ThemePptTaskPane.vb  # Docmee task pane for PPT generation
└── ThisAddIn.vb         # PowerPoint add-in entry point
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| PowerPoint ribbon customization | Ribbon1.vb | PowerPoint tab and button definitions |
| Docmee PPT generation | ThemePptTaskPane.vb | Task pane and WebView integration |
| Slide operations | Ribbon1.vb | Current-slide AI creation and replacement |
| Add-in initialization | ThisAddIn.vb | Startup and shutdown logic |

## CONVENTIONS
- PowerPoint-specific functionality only
- Uses PowerPoint interop through ShareRibbon services
- Follows PowerPoint VSTO patterns

## ANTI-PATTERNS
- Never access Excel or Word objects
- Never duplicate shared functionality from ShareRibbon
- Never bypass PowerPoint slide protection
