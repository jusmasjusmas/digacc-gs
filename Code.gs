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
  issues = issues.concat(manualCheck("document_language", -1,
    "Document language should be specified",
    "Verify the presentation language is set correctly in File > Language. Cannot be checked automatically."));

  // Slides
  issues = issues.concat(checkSlideTitles(slides, apiKey));
  issues = issues.concat(checkDuplicateTitles(slides));
  issues = issues.concat(checkEmptySlides(slides));

  // Tables
  issues = issues.concat(checkTableAltText(slides));
  issues = issues.concat(manualCheck("merged_cells", -1,
    "The use of merged cells is not recommended",
    "Check your tables for merged cells. Merged cells break screen reader navigation. Cannot be detected automatically."));
  issues = issues.concat(checkEmptyCells(slides));

  // Elements
  issues = issues.concat(checkImageAltText(slides, apiKey));
  issues = issues.concat(checkElementAltText(slides));
  issues = issues.concat(checkEmptyTextBoxes(slides));
  issues = issues.concat(manualCheck("broken_lists", -1,
    "Lists should not be broken apart",
    "Check that bullet/numbered lists are not split by blank lines or formatting breaks. Cannot be detected automatically."));

  // Contents
  issues = issues.concat(checkSmallText(slides));
  issues = issues.concat(checkColorContrast(slides));
  issues = issues.concat(manualCheck("inline_style", -1,
    "In-line style changes may lack clear meaning",
    "Check that mid-paragraph font or color changes are meaningful and not purely decorative. Cannot be detected automatically."));
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
    title:          "Presentation has no meaningful title",
    description:    "The file name is the presentation title. Rename the file to something descriptive.",
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
        title:          "Slide " + (i + 1) + " has no title placeholder",
        description:    "Screen readers identify slides by their title (WCAG 2.4.6).",
        suggestedValue: suggested,
        autoFixable:    suggested !== "Slide " + (i + 1),
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
          title:          "Slide " + (i + 1) + " has an empty title",
          description:    "Title placeholder exists but has no text.",
          suggestedValue: suggested,
          autoFixable:    suggested !== "Slide " + (i + 1),
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
        title:          "Slide " + (i + 1) + ": duplicate title \u201c" + text + "\u201d",
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
        title:       "Slide " + (i + 1) + " appears to be empty",
        description: "Empty slides provide no content for screen reader users.",
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
        title:          "Slide " + (i + 1) + ": table missing description",
        description:    "Tables need alt text to summarise their content (WCAG 1.1.1).",
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
        title:          "Slide " + (i + 1) + ": table has " + count + " empty cell(s)",
        description:    "Empty cells should contain text or a placeholder so screen readers can announce them.",
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
        title:          "Slide " + (i + 1) + ": image missing alt text",
        description:    "Images need alternative text for screen readers (WCAG 1.1.1).",
        suggestedValue: suggested,
        autoFixable:    !!apiKey,
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
        title:          "Slide " + (i + 1) + ": " + label.toLowerCase() + " missing alt text",
        description:    label + "s need alternative text for screen readers (WCAG 1.1.1).",
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
      if (ph !== SlidesApp.PlaceholderType.NONE &&
          ph !== SlidesApp.PlaceholderType.NOT_PLACEHOLDER) return;
      if (shape.getText().asString().trim()) return;
      issues.push(newIssue({
        checkType:   "empty_textbox",
        slideIndex:  i,
        elementId:   el.getObjectId(),
        severity:    "warning",
        title:       "Slide " + (i + 1) + ": empty text box",
        description: "Empty text boxes are announced as blank elements by screen readers.",
        autoFixable: true,
      }));
    });
  });
  return issues;
}

function checkSmallText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      try {
        el.asShape().getText().getParagraphs().forEach(function(para) {
          para.getRange().getTextRuns().forEach(function(run) {
            if (!run.asString().trim()) return;
            var size = run.getTextStyle().getFontSize();
            if (size && size > 0 && size < FINE_PRINT_PT) {
              issues.push(newIssue({
                checkType:   "small_text",
                slideIndex:  i,
                elementId:   el.getObjectId(),
                severity:    "warning",
                title:       "Slide " + (i + 1) + ": " + Math.round(size) + "pt text",
                description: "Text is " + Math.round(size) + "pt \u2014 increase to \u2265" + FINE_PRINT_PT + "pt unless intentional.",
                autoFixable: false,
              }));
            }
          });
        });
      } catch(e) {}
    });
  });
  return issues;
}

function checkColorContrast(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      try {
        el.asShape().getText().getParagraphs().forEach(function(para) {
          para.getRange().getTextRuns().forEach(function(run) {
            if (!run.asString().trim()) return;
            var color = run.getTextStyle().getForegroundColor();
            if (!color || color.getColorType() !== SlidesApp.ColorType.RGB) return;
            var rgb = color.asRgbColor();
            if (isTooLight(rgb.getRed(), rgb.getGreen(), rgb.getBlue())) {
              issues.push(newIssue({
                checkType:   "low_contrast",
                slideIndex:  i,
                elementId:   el.getObjectId(),
                severity:    "error",
                title:       "Slide " + (i + 1) + ": low-contrast text (#" + toHex(rgb) + ")",
                description: "This text color may fail the WCAG AA contrast ratio of 4.5:1.",
                autoFixable: false,
              }));
            }
          });
        });
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
        if (paras.length < 2) return;
        var last = paras[paras.length - 1].getRange().asString();
        var prev = paras[paras.length - 2].getRange().asString();
        if (!last.trim() && !prev.trim()) {
          issues.push(newIssue({
            checkType:   "trailing_lines",
            slideIndex:  i,
            elementId:   el.getObjectId(),
            severity:    "info",
            title:       "Slide " + (i + 1) + ": trailing empty lines",
            description: "Trailing empty paragraphs are read aloud by screen readers as blank content.",
            autoFixable: false,
          }));
        }
      } catch(e) {}
    });
  });
  return issues;
}

// Helper: add a manual-review-only issue for checks that can't be automated
function manualCheck(checkType, slideIndex, title, description) {
  return [newIssue({
    checkType:   checkType,
    slideIndex:  slideIndex,
    severity:    "info",
    title:       title,
    description: description,
    autoFixable: false,
    manual:      true,
  })];
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
        if (el) el.remove();
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

      // human-review: just mark as acknowledged
      case "small_text":
      case "low_contrast":
      case "empty_slide":
      case "trailing_lines":
      case "document_language":
      case "merged_cells":
      case "broken_lists":
      case "inline_style":
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
