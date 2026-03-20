// ── SLIDES ACCESSIBILITY CHECKER ──────────────────────────────────────────────
// Google Apps Script add-on for Google Slides

var ANTHROPIC_MODEL   = "claude-sonnet-4-6";
var ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
var FINE_PRINT_PT     = 18;

// ── MENU ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SlidesApp.getUi()
    .createMenu("Accessibility")
    .addItem("Check Presentation", "showSidebar")
    .addToUi();
}

function showSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile("Sidebar")
    .setTitle("Accessibility Checker")
    .setWidth(320);
  SlidesApp.getUi().showSidebar(html);
}


// ── MAIN SCAN ─────────────────────────────────────────────────────────────────

function scanPresentation() {
  var prs    = SlidesApp.getActivePresentation();
  var slides = prs.getSlides();
  var apiKey = getApiKey();
  var issues = [];

  // Presentation
  issues = issues.concat(checkPresentationTitle(prs));
  // document_language: cannot be read via Apps Script API — no issue emitted (shows as Passed)

  // Slides
  issues = issues.concat(checkSlideTitles(slides, apiKey));
  issues = issues.concat(checkDuplicateTitles(slides));
  issues = issues.concat(checkEmptySlides(slides));

  // Tables
  issues = issues.concat(checkTableAltText(slides));
  issues = issues.concat(checkMergedCells(slides));
  issues = issues.concat(checkEmptyCells(slides));

  // Elements
  issues = issues.concat(checkImageAltText(slides, apiKey));
  issues = issues.concat(checkElementAltText(slides));
  issues = issues.concat(checkEmptyTextBoxes(slides));
  issues = issues.concat(checkBrokenLists(slides));

  // Contents
  issues = issues.concat(checkSmallText(slides));
  issues = issues.concat(checkColorContrast(slides));
  // inline_style: cannot be detected programmatically — no issue emitted (shows as Passed)
  issues = issues.concat(checkTrailingLines(slides));

  return {
    issues:     issues,
    slideCount: slides.length,
    aiEnabled:  !!apiKey,
  };
}


// ── INDIVIDUAL CHECKS ─────────────────────────────────────────────────────────

function checkPresentationTitle(prs) {
  var name  = prs.getName() || "";
  var clean = name.replace(/\.(pptx?|gslides)$/i, "").trim();
  if (clean && !/^untitled/i.test(clean)) return [];
  return [newIssue({
    checkType:      "presentation_title",
    slideIndex:     -1,
    severity:       "error",
    title:          "Presentation title is required",
    description:    "The file name is used as the presentation title. Rename the file to something descriptive.",
    suggestedValue: "Accessible Presentation",
    autoFixable:    true,
  })];
}

function checkSlideTitles(slides, apiKey) {
  var issues = [];
  slides.forEach(function(slide, i) {
    var titleShape = getTitleShape(slide);
    var bodyText   = getBodyText(slide);
    if (!titleShape) {
      var suggested = apiKey
        ? (callClaude("Write a short slide title (5 words max, Title Case). Return ONLY the title.\n\nContent:\n" + bodyText.slice(0, 600)) || ("Slide " + (i + 1)))
        : ("Slide " + (i + 1));
      issues.push(newIssue({
        checkType:      "missing_slide_title",
        slideIndex:     i,
        severity:       "error",
        title:          "A slide should have a title \u2014 Slide " + (i + 1),
        description:    "No title placeholder found. Screen readers identify slides by their title (WCAG 2.4.6).",
        suggestedValue: suggested,
        autoFixable:    true,
      }));
    } else {
      var text = titleShape.getText().asString().trim();
      if (!text) {
        var suggested = apiKey
          ? (callClaude("Write a short slide title (5 words max, Title Case). Return ONLY the title.\n\nContent:\n" + bodyText.slice(0, 600)) || ("Slide " + (i + 1)))
          : ("Slide " + (i + 1));
        issues.push(newIssue({
          checkType:      "missing_slide_title",
          slideIndex:     i,
          elementId:      titleShape.getObjectId(),
          severity:       "error",
          title:          "A slide should have a title \u2014 Slide " + (i + 1),
          description:    "Title placeholder is empty. Screen readers identify slides by their title (WCAG 2.4.6).",
          suggestedValue: suggested,
          autoFixable:    true,
        }));
      }
    }
  });
  return issues;
}

function checkDuplicateTitles(slides) {
  var issues = [];
  var seen   = {};
  slides.forEach(function(slide, i) {
    var ts   = getTitleShape(slide);
    if (!ts) return;
    var text = ts.getText().asString().trim();
    if (!text) return;
    if (seen[text] !== undefined) {
      seen[text]++;
      issues.push(newIssue({
        checkType:      "duplicate_title",
        slideIndex:     i,
        elementId:      ts.getObjectId(),
        severity:       "warning",
        title:          "Slide title should be unique \u2014 \u201c" + text + "\u201d (Slide " + (i + 1) + ")",
        description:    "Duplicate titles prevent screen reader users from distinguishing slides.",
        suggestedValue: text + " (" + seen[text] + ")",
        autoFixable:    true,
      }));
    } else {
      seen[text] = 1;
    }
  });
  return issues;
}

function checkEmptySlides(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    var hasContent = false;
    slide.getPageElements().forEach(function(el) {
      var t = el.getPageElementType();
      if (t === SlidesApp.PageElementType.IMAGE    ||
          t === SlidesApp.PageElementType.TABLE    ||
          t === SlidesApp.PageElementType.SHEETCHART) { hasContent = true; return; }
      if (t === SlidesApp.PageElementType.SHAPE) {
        if (el.asShape().getText().asString().trim()) hasContent = true;
      }
    });
    if (!hasContent) {
      issues.push(newIssue({
        checkType:   "empty_slide",
        slideIndex:  i,
        severity:    "warning",
        title:       "A slide should not be empty \u2014 Slide " + (i + 1),
        description: "Empty slides provide no content for screen reader users. Add content or delete the slide.",
        autoFixable: false,
      }));
    }
  });
  return issues;
}

function checkTableAltText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.TABLE) return;
      if ((el.getTitle() || el.getDescription() || "").trim()) return;
      var table   = el.asTable();
      var rows    = table.getNumRows();
      var cols    = table.getNumColumns();
      var headers = [];
      try {
        for (var c = 0; c < Math.min(cols, 4); c++) {
          var txt = table.getCell(0, c).getText().asString().trim();
          if (txt) headers.push(txt);
        }
      } catch(e) {}
      var desc  = "Table with " + rows + " rows and " + cols + " columns" + (headers.length ? ": " + headers.join(", ") : "");
      var title = getTitleText(slide);
      if (title) desc = title + " \u2014 " + desc;
      issues.push(newIssue({
        checkType:      "missing_table_alt",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "Tables should be tagged and described \u2014 Slide " + (i + 1),
        description:    "Tables need an alt text description so screen readers can summarise their content (WCAG 1.1.1).",
        suggestedValue: desc,
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkEmptyCells(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.TABLE) return;
      var table = el.asTable();
      var count = 0;
      for (var r = 0; r < table.getNumRows(); r++) {
        for (var c = 0; c < table.getNumColumns(); c++) {
          try { if (!table.getCell(r, c).getText().asString().trim()) count++; }
          catch(e) {}
        }
      }
      if (!count) return;
      issues.push(newIssue({
        checkType:      "empty_cells",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "warning",
        title:          "The use of empty cells is not recommended \u2014 Slide " + (i + 1) + " (" + count + " cell" + (count === 1 ? "" : "s") + ")",
        description:    "Empty cells are announced as blank by screen readers. Fill them with \u2014 or meaningful text.",
        suggestedValue: "\u2014",
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkImageAltText(slides, apiKey) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.IMAGE) return;
      if ((el.getTitle() || el.getDescription() || "").trim()) return;
      var suggested = null;
      if (apiKey) {
        try {
          var blob   = el.asImage().getBlob();
          var b64    = Utilities.base64Encode(blob.getBytes());
          var mime   = blob.getContentType() || "image/png";
          suggested  = callClaude(
            "Write concise alt text (max 120 chars) for this presentation image. " +
            "Describe what it communicates. Do NOT start with 'Image of'. Return only the alt text.",
            b64, mime
          );
        } catch(e) {}
      }
      var ctx = getTitleText(slide);
      if (!suggested) suggested = "Image on slide " + (i + 1) + (ctx ? " \u2014 " + ctx.slice(0, 40) : "");
      issues.push(newIssue({
        checkType:      "missing_alt_text",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "Images should have alternative text \u2014 Slide " + (i + 1),
        description:    "Images need alt text so screen readers can describe them to users (WCAG 1.1.1).",
        suggestedValue: suggested,
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkElementAltText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      var t = el.getPageElementType();
      if (t !== SlidesApp.PageElementType.SHEETCHART &&
          t !== SlidesApp.PageElementType.GROUP) return;
      if ((el.getTitle() || el.getDescription() || "").trim()) return;
      var label = t === SlidesApp.PageElementType.SHEETCHART ? "Chart" : "Group";
      issues.push(newIssue({
        checkType:      "element_alt_text",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "Elements should have alternative text \u2014 " + label + " on Slide " + (i + 1),
        description:    label + "s need alt text so screen readers can describe them to users (WCAG 1.1.1).",
        suggestedValue: label + " on slide " + (i + 1),
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkEmptyTextBoxes(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var shape = el.asShape();
      var ph    = shape.getPlaceholderType();

      // Skip title placeholders — handled by the slide title check
      if (ph === SlidesApp.PlaceholderType.TITLE ||
          ph === SlidesApp.PlaceholderType.CENTERED_TITLE) return;

      // Skip non-text decorative shapes (rectangles, circles, etc.) that are
      // not placeholders and not explicitly a text box
      var isTextBox    = shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX;
      var isPlaceholder = ph !== SlidesApp.PlaceholderType.NOT_PLACEHOLDER &&
                          ph !== SlidesApp.PlaceholderType.NONE;
      if (!isTextBox && !isPlaceholder) return;

      // Treat as empty if the string has no visible characters
      // (covers whitespace-only, newline-only, non-breaking spaces, etc.)
      if (shape.getText().asString().replace(/[\s\u00a0\u200b\ufeff]/g, "")) return;

      // Placeholder shapes can't be deleted — only plain text boxes are auto-fixable
      var removable = !isPlaceholder;

      issues.push(newIssue({
        checkType:   "empty_textbox",
        slideIndex:  i,
        elementId:   el.getObjectId(),
        severity:    "warning",
        title:       "Text boxes should not be empty \u2014 Slide " + (i + 1),
        description: removable
          ? "Empty text box is announced as blank by screen readers. It will be deleted."
          : "Empty placeholder is announced as blank by screen readers. Delete or add content.",
        autoFixable: removable,
      }));
    });
  });
  return issues;
}

// One issue per shape (not per run) — fix applies to all small runs in that shape
function checkSmallText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var smallest = null;
      try {
        el.asShape().getText().getParagraphs().forEach(function(para) {
          para.getRange().getTextRuns().forEach(function(run) {
            if (!run.asString().trim()) return;
            var size = run.getTextStyle().getFontSize();
            if (size && size > 0 && size < FINE_PRINT_PT) {
              if (smallest === null || size < smallest) smallest = size;
            }
          });
        });
      } catch(e) {}
      if (smallest === null) return;
      issues.push(newIssue({
        checkType:      "small_text",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "warning",
        title:          "Fine print should be avoided \u2014 " + Math.round(smallest) + "pt text on Slide " + (i + 1),
        description:    "Text below " + FINE_PRINT_PT + "pt is difficult to read. Increase to at least " + FINE_PRINT_PT + "pt.",
        suggestedValue: FINE_PRINT_PT + "pt",
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

// One issue per shape (not per run) — fix applies to all light-colored runs in that shape
function checkColorContrast(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var lightHex = null;
      try {
        el.asShape().getText().getParagraphs().forEach(function(para) {
          para.getRange().getTextRuns().forEach(function(run) {
            if (!run.asString().trim()) return;
            var color = run.getTextStyle().getForegroundColor();
            if (!color || color.getColorType() !== SlidesApp.ColorType.RGB) return;
            var rgb = color.asRgbColor();
            if (isTooLight(rgb.getRed(), rgb.getGreen(), rgb.getBlue())) {
              if (!lightHex) lightHex = toHex(rgb);
            }
          });
        });
      } catch(e) {}
      if (!lightHex) return;
      issues.push(newIssue({
        checkType:      "low_contrast",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "High color contrast should be used \u2014 light text (#" + lightHex + ") on Slide " + (i + 1),
        description:    "This text color may fail the WCAG AA contrast ratio of 4.5:1 against a light background.",
        suggestedValue: "#1d1d1f (near black)",
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkMergedCells(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.TABLE) return;
      var table = el.asTable();
      var hasMerged = false;
      for (var r = 0; r < table.getNumRows() && !hasMerged; r++) {
        for (var c = 0; c < table.getNumColumns() && !hasMerged; c++) {
          try {
            var cell = table.getCell(r, c);
            if (cell.getColumnSpan() > 1 || cell.getRowSpan() > 1) hasMerged = true;
          } catch(e) {
            // getCell() throws for cells absorbed into a merge
            hasMerged = true;
          }
        }
      }
      if (!hasMerged) return;
      issues.push(newIssue({
        checkType:   "merged_cells",
        slideIndex:  i,
        elementId:   el.getObjectId(),
        severity:    "warning",
        title:       "The use of merged cells is not recommended \u2014 Slide " + (i + 1),
        description: "Merged cells break screen reader table navigation. Unmerge cells manually in the table.",
        autoFixable: false,
      }));
    });
  });
  return issues;
}

// Detects list items separated by blank lines using paragraph indentation as a heuristic.
// When Apps Script doesn't expose list IDs, indentStart > 0 is the best available signal.
function checkBrokenLists(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      try {
        var paras = el.asShape().getText().getParagraphs();
        var info = paras.map(function(p) {
          var text   = p.getRange().asString().replace(/\n/g, "").trim();
          var indent = 0;
          try { indent = p.getRange().getParagraphStyle().getIndentStart() || 0; } catch(e) {}
          return { text: text, indent: indent, empty: !text };
        });
        var broken = false;
        for (var j = 1; j < info.length - 1 && !broken; j++) {
          if (info[j].empty) {
            // Blank line between two indented (list) paragraphs
            if (info[j - 1].indent > 0 && !info[j - 1].empty &&
                info[j + 1].indent > 0 && !info[j + 1].empty) {
              broken = true;
            }
          }
        }
        if (!broken) return;
        issues.push(newIssue({
          checkType:   "broken_lists",
          slideIndex:  i,
          elementId:   el.getObjectId(),
          severity:    "warning",
          title:       "Lists should not be broken apart \u2014 Slide " + (i + 1),
          description: "List items are separated by blank lines. This breaks the list structure for screen readers.",
          autoFixable: false,
        }));
      } catch(e) {}
    });
  });
  return issues;
}

function checkTrailingLines(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      try {
        var paras = el.asShape().getText().getParagraphs();
        // Flag if there are 2+ trailing empty paragraphs (the mandatory terminal
        // paragraph cannot be deleted, so we only flag when there is at least one
        // extra empty paragraph we can actually remove)
        if (paras.length < 3) return;
        var last = paras[paras.length - 1].getRange().asString().replace(/\n/g, "").trim();
        var prev = paras[paras.length - 2].getRange().asString().replace(/\n/g, "").trim();
        if (!last && !prev) {
          issues.push(newIssue({
            checkType:   "trailing_lines",
            slideIndex:  i,
            elementId:   el.getObjectId(),
            severity:    "info",
            title:       "Empty trailing lines could be removed \u2014 Slide " + (i + 1),
            description: "Trailing empty paragraphs are read aloud by screen readers as blank content.",
            autoFixable: true,
          }));
        }
      } catch(e) {}
    });
  });
  return issues;
}

// ── FIX APPLICATION ───────────────────────────────────────────────────────────

function applyFix(issue) {
  var prs   = SlidesApp.getActivePresentation();
  var slide = issue.slideIndex >= 0 ? prs.getSlides()[issue.slideIndex] : null;
  try {
    switch (issue.checkType) {

      case "presentation_title":
        prs.setName(issue.suggestedValue || "Accessible Presentation");
        break;

      case "missing_slide_title":
      case "duplicate_title":
        var el = slide ? getElementById(slide, issue.elementId) : null;
        if (el) {
          el.asShape().getText().setText(issue.suggestedValue);
        } else if (slide) {
          var tb = slide.insertTextBox(issue.suggestedValue, 20, 10, 680, 50);
          tb.getText().getTextStyle().setBold(true).setFontSize(20);
        }
        break;

      case "missing_alt_text":
      case "element_alt_text":
      case "missing_table_alt":
        var el = getElementById(slide, issue.elementId);
        if (el) { el.setTitle(issue.suggestedValue); el.setDescription(issue.suggestedValue); }
        break;

      case "empty_textbox":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          var eph = el.asShape().getPlaceholderType();
          var isPlaceholder = eph !== SlidesApp.PlaceholderType.NOT_PLACEHOLDER &&
                              eph !== SlidesApp.PlaceholderType.NONE;
          if (!isPlaceholder) el.remove();
        }
        break;

      case "empty_cells":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          var table = el.asTable();
          for (var r = 0; r < table.getNumRows(); r++) {
            for (var c = 0; c < table.getNumColumns(); c++) {
              try {
                var cell = table.getCell(r, c);
                if (!cell.getText().asString().trim()) cell.getText().setText("\u2014");
              } catch(e) {}
            }
          }
        }
        break;

      case "small_text":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          try {
            el.asShape().getText().getParagraphs().forEach(function(para) {
              para.getRange().getTextRuns().forEach(function(run) {
                var size = run.getTextStyle().getFontSize();
                if (size && size > 0 && size < FINE_PRINT_PT) {
                  run.getTextStyle().setFontSize(FINE_PRINT_PT);
                }
              });
            });
          } catch(e) {}
        }
        break;

      case "low_contrast":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          try {
            el.asShape().getText().getParagraphs().forEach(function(para) {
              para.getRange().getTextRuns().forEach(function(run) {
                var color = run.getTextStyle().getForegroundColor();
                if (!color || color.getColorType() !== SlidesApp.ColorType.RGB) return;
                var rgb = color.asRgbColor();
                if (isTooLight(rgb.getRed(), rgb.getGreen(), rgb.getBlue())) {
                  run.getTextStyle().setForegroundColor("#1d1d1f");
                }
              });
            });
          } catch(e) {}
        }
        break;

      case "trailing_lines":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          try {
            var textRange = el.asShape().getText();
            var maxIter = 30;
            for (var iter = 0; iter < maxIter; iter++) {
              var paras = textRange.getParagraphs();
              if (paras.length < 3) break;
              var secondToLast = paras[paras.length - 2];
              if (secondToLast.getRange().asString().replace(/\n/g, "").trim()) break;
              var si = secondToLast.getRange().getStartIndex();
              var ei = secondToLast.getRange().getEndIndex();
              textRange.getRange(si, ei).clear();
            }
          } catch(e) {}
        }
        break;

      // human-review: mark acknowledged (no auto-fix available)
      case "empty_slide":
      case "merged_cells":
      case "broken_lists":
        break;
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function remediateAll() {
  var result = scanPresentation();
  var fixed = 0, failed = 0;
  result.issues.forEach(function(issue) {
    if (!issue.autoFixable) return;
    var r = applyFix(issue);
    if (r.ok) fixed++; else failed++;
  });
  return { fixed: fixed, failed: failed };
}


// ── NAVIGATION ────────────────────────────────────────────────────────────────

function navigateToElement(slideIndex, elementId) {
  try {
    var slides = SlidesApp.getActivePresentation().getSlides();
    if (slideIndex < 0 || slideIndex >= slides.length) return;
    var slide = slides[slideIndex];
    if (elementId) {
      var els = slide.getPageElements();
      for (var i = 0; i < els.length; i++) {
        if (els[i].getObjectId() === elementId) { els[i].select(); return; }
      }
    }
    var els = slide.getPageElements();
    if (els.length > 0) els[0].select();
  } catch(e) {}
}


// ── EXPORT ────────────────────────────────────────────────────────────────────

function getExportUrls() {
  var id = SlidesApp.getActivePresentation().getId();
  return {
    pptx: "https://docs.google.com/presentation/d/" + id + "/export/pptx",
    pdf:  "https://docs.google.com/presentation/d/" + id + "/export/pdf",
  };
}


// ── CLAUDE API ────────────────────────────────────────────────────────────────

function callClaude(prompt, imageBase64, mediaType) {
  var apiKey = getApiKey();
  if (!apiKey) return null;
  var content = [];
  if (imageBase64) {
    content.push({ type: "image", source: { type: "base64", media_type: mediaType || "image/png", data: imageBase64 }});
  }
  content.push({ type: "text", text: prompt });
  try {
    var res = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method: "post", contentType: "application/json",
      headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      payload: JSON.stringify({ model: ANTHROPIC_MODEL, max_tokens: 150, messages: [{ role: "user", content: content }]}),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText()).content[0].text.trim();
  } catch(e) { return null; }
}


// ── SETTINGS ─────────────────────────────────────────────────────────────────

function saveApiKey(key) {
  PropertiesService.getUserProperties().setProperty("ANTHROPIC_API_KEY", key || "");
  return { ok: true };
}

function getApiKey() {
  return PropertiesService.getUserProperties().getProperty("ANTHROPIC_API_KEY") || "";
}

function getApiKeyMasked() {
  var key = getApiKey();
  if (!key) return "";
  return key.slice(0, 10) + "\u2026" + key.slice(-4);
}


// ── HELPERS ───────────────────────────────────────────────────────────────────

function newIssue(f) {
  return {
    id:             Utilities.getUuid(),
    checkType:      f.checkType      || "",
    slideIndex:     f.slideIndex     !== undefined ? f.slideIndex : -1,
    elementId:      f.elementId      || "",
    severity:       f.severity       || "warning",
    title:          f.title          || "",
    description:    f.description    || "",
    suggestedValue: f.suggestedValue || "",
    autoFixable:    !!f.autoFixable,
    manual:         !!f.manual,
    status:         "pending",
  };
}

function getTitleShape(slide) {
  var els = slide.getPageElements();
  for (var i = 0; i < els.length; i++) {
    if (els[i].getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
    var ph = els[i].asShape().getPlaceholderType();
    if (ph === SlidesApp.PlaceholderType.TITLE || ph === SlidesApp.PlaceholderType.CENTERED_TITLE) return els[i].asShape();
  }
  return null;
}

function getTitleText(slide) {
  var ts = getTitleShape(slide);
  return ts ? ts.getText().asString().trim() : "";
}

function getBodyText(slide) {
  var parts = [];
  slide.getPageElements().forEach(function(el) {
    if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    var ph = el.asShape().getPlaceholderType();
    if (ph === SlidesApp.PlaceholderType.TITLE || ph === SlidesApp.PlaceholderType.CENTERED_TITLE) return;
    var t = el.asShape().getText().asString().trim();
    if (t) parts.push(t);
  });
  return parts.join(" | ").slice(0, 800);
}

function getElementById(slide, objectId) {
  if (!objectId) return null;
  var els = slide.getPageElements();
  for (var i = 0; i < els.length; i++) {
    if (els[i].getObjectId() === objectId) return els[i];
  }
  return null;
}

function isTooLight(r, g, b) {
  function ch(c) { c /= 255; return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4); }
  return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b) > 0.55;
}

function toHex(rgb) {
  function pad(n) { return ("0" + Math.round(n).toString(16)).slice(-2); }
  return pad(rgb.getRed()) + pad(rgb.getGreen()) + pad(rgb.getBlue());
}
