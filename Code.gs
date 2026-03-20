// ── SLIDES ACCESSIBILITY CHECKER ──────────────────────────────────────────────
// Google Apps Script add-on for Google Slides
// Checks for and fixes WCAG 2.1 AA / Title II accessibility issues.

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
  var prs     = SlidesApp.getActivePresentation();
  var slides  = prs.getSlides();
  var apiKey  = getApiKey();
  var issues  = [];

  issues = issues.concat(checkPresentationTitle(prs));
  issues = issues.concat(checkSlideTitles(slides, apiKey));
  issues = issues.concat(checkDuplicateTitles(slides));
  issues = issues.concat(checkImageAltText(slides, apiKey));
  issues = issues.concat(checkTableAltText(slides));
  issues = issues.concat(checkEmptyTextBoxes(slides));
  issues = issues.concat(checkSpeakerNotes(slides));
  issues = issues.concat(checkSmallText(slides));
  issues = issues.concat(checkColorContrast(slides));

  var errors   = issues.filter(function(i) { return i.severity === "error"; }).length;
  var warnings = issues.filter(function(i) { return i.severity === "warning"; }).length;

  return {
    issues:     issues,
    errors:     errors,
    warnings:   warnings,
    slideCount: slides.length,
    aiEnabled:  !!apiKey,
  };
}


// ── INDIVIDUAL CHECKS ─────────────────────────────────────────────────────────

function checkPresentationTitle(prs) {
  var name  = prs.getName() || "";
  var clean = name.replace(/\.(pptx?|gslides)$/i, "").trim();
  var bad   = !clean || /^untitled/i.test(clean);
  if (!bad) return [];
  return [newIssue({
    checkType:      "presentation_title",
    slideIndex:     -1,
    severity:       "error",
    title:          "Missing presentation title",
    description:    "The file has no meaningful title. Screen readers announce the presentation title when it opens.",
    suggestedValue: "Untitled Presentation",
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
        ? (callClaude("Write a short slide title (5 words max, Title Case). Return ONLY the title.\n\nContent:\n" + bodyText.slice(0, 600), null, null) || ("Slide " + (i + 1)))
        : ("Slide " + (i + 1));
      var isPlaceholder = suggested === "Slide " + (i + 1);
      issues.push(newIssue({
        checkType:      "missing_slide_title",
        slideIndex:     i,
        severity:       "error",
        title:          "Slide " + (i + 1) + ": No title placeholder",
        description:    "Slides need a title for screen readers to identify them (WCAG 2.4.6).",
        suggestedValue: suggested,
        autoFixable:    !isPlaceholder,
      }));
    } else {
      var text = titleShape.getText().asString().trim();
      if (!text) {
        var suggested = apiKey
          ? (callClaude("Write a short slide title (5 words max, Title Case). Return ONLY the title.\n\nContent:\n" + bodyText.slice(0, 600), null, null) || ("Slide " + (i + 1)))
          : ("Slide " + (i + 1));
        var isPlaceholder = suggested === "Slide " + (i + 1);
        issues.push(newIssue({
          checkType:      "missing_slide_title",
          slideIndex:     i,
          elementId:      titleShape.getObjectId(),
          severity:       "error",
          title:          "Slide " + (i + 1) + ": Empty slide title",
          description:    "Title placeholder exists but contains no text.",
          suggestedValue: suggested,
          autoFixable:    !isPlaceholder,
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
      var deduped = text + " (" + seen[text] + ")";
      issues.push(newIssue({
        checkType:      "duplicate_title",
        slideIndex:     i,
        elementId:      ts.getObjectId(),
        severity:       "warning",
        title:          "Slide " + (i + 1) + ": Duplicate title \u201c" + text + "\u201d",
        description:    "Duplicate titles prevent screen reader users from telling slides apart.",
        suggestedValue: deduped,
        autoFixable:    true,
      }));
    } else {
      seen[text] = 1;
    }
  });
  return issues;
}

function checkImageAltText(slides, apiKey) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.IMAGE) return;
      var existing = (el.getTitle() || el.getDescription() || "").trim();
      if (existing) return;

      var suggested = null;
      if (apiKey) {
        try {
          var blob     = el.asImage().getBlob();
          var base64   = Utilities.base64Encode(blob.getBytes());
          var mimeType = blob.getContentType() || "image/png";
          suggested = callClaude(
            "Write concise alt text (max 120 chars) for this presentation image. " +
            "Describe what it communicates. Do NOT start with 'Image of'. Return only the alt text.",
            base64, mimeType
          );
        } catch(e) { /* blob unavailable */ }
      }
      var isPlaceholder = !suggested;
      if (!suggested) {
        var context = getTitleText(slide);
        suggested = "Image on slide " + (i + 1) + (context ? " \u2014 " + context.slice(0, 40) : "");
      }

      issues.push(newIssue({
        checkType:      "missing_alt_text",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "Slide " + (i + 1) + ": Image missing alt text",
        description:    "Images need alternative text for screen readers (WCAG 1.1.1).",
        suggestedValue: suggested,
        autoFixable:    !isPlaceholder,
      }));
    });
  });
  return issues;
}

function checkTableAltText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.TABLE) return;
      var existing = (el.getTitle() || el.getDescription() || "").trim();
      if (existing) return;
      var table    = el.asTable();
      var rows     = table.getNumRows();
      var cols     = table.getNumColumns();
      var headers  = [];
      try {
        for (var c = 0; c < Math.min(cols, 4); c++) {
          var txt = table.getCell(0, c).getText().asString().trim();
          if (txt) headers.push(txt);
        }
      } catch(e) {}
      var desc = "Table with " + rows + " rows and " + cols + " columns" +
                 (headers.length ? ": " + headers.join(", ") : "");
      var slideTitle = getTitleText(slide);
      if (slideTitle) desc = slideTitle + " \u2014 " + desc;

      issues.push(newIssue({
        checkType:      "missing_table_alt",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "error",
        title:          "Slide " + (i + 1) + ": Table missing description",
        description:    "Tables need alt text to summarise their content (WCAG 1.1.1).",
        suggestedValue: desc,
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
      // Skip title / subtitle / body placeholders
      var phType = shape.getPlaceholderType();
      if (phType !== SlidesApp.PlaceholderType.NONE &&
          phType !== SlidesApp.PlaceholderType.NOT_PLACEHOLDER) return;
      var text = shape.getText().asString().trim();
      if (text) return;
      issues.push(newIssue({
        checkType:      "empty_textbox",
        slideIndex:     i,
        elementId:      el.getObjectId(),
        severity:       "warning",
        title:          "Slide " + (i + 1) + ": Empty text box",
        description:    "Empty text boxes are announced as blank elements by screen readers.",
        suggestedValue: "",
        autoFixable:    true,
      }));
    });
  });
  return issues;
}

function checkSpeakerNotes(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    try {
      var notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
      if (!notes) {
        issues.push(newIssue({
          checkType:      "no_speaker_notes",
          slideIndex:     i,
          severity:       "warning",
          title:          "Slide " + (i + 1) + ": No speaker notes",
          description:    "Speaker notes provide context for screen reader users and document accessibility.",
          suggestedValue: "",
          autoFixable:    false,
        }));
      }
    } catch(e) {}
  });
  return issues;
}

function checkSmallText(slides) {
  var issues = [];
  slides.forEach(function(slide, i) {
    slide.getPageElements().forEach(function(el) {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      var shape = el.asShape();
      try {
        var textRange = shape.getText();
        textRange.getParagraphs().forEach(function(para) {
          para.getRange().getTextRuns().forEach(function(run) {
            var style = run.getTextStyle();
            var size  = style.getFontSize();
            if (size && size > 0 && size < FINE_PRINT_PT) {
              issues.push(newIssue({
                checkType:      "small_text",
                slideIndex:     i,
                elementId:      el.getObjectId(),
                severity:       "warning",
                title:          "Slide " + (i + 1) + ": " + Math.round(size) + "pt text in \u201c" + (el.getTitle() || shape.getPlaceholderType() || "shape") + "\u201d",
                description:    "Text is " + Math.round(size) + "pt \u2014 increase to \u2265" + FINE_PRINT_PT + "pt unless intentional.",
                suggestedValue: "",
                autoFixable:    false,
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
        var shape = el.asShape();
        var runs  = shape.getText().getParagraphs()
                         .reduce(function(acc, p) {
                           return acc.concat(p.getRange().getTextRuns());
                         }, []);
        runs.forEach(function(run) {
          if (!run.asString().trim()) return;
          var color = run.getTextStyle().getForegroundColor();
          if (!color || color.getColorType() !== SlidesApp.ColorType.RGB) return;
          var rgb = color.asRgbColor();
          if (isTooLight(rgb.getRed(), rgb.getGreen(), rgb.getBlue())) {
            issues.push(newIssue({
              checkType:      "low_contrast",
              slideIndex:     i,
              elementId:      el.getObjectId(),
              severity:       "error",
              title:          "Slide " + (i + 1) + ": Low-contrast text",
              description:    "Text color (#" + toHex(rgb) + ") may fail WCAG AA contrast ratio (4.5:1).",
              suggestedValue: "",
              autoFixable:    false,
            }));
          }
        });
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
          // No title placeholder — add one as a text box at the top
          var tb = slide.insertTextBox(issue.suggestedValue, 20, 10, 680, 50);
          tb.getText().getTextStyle().setBold(true).setFontSize(20);
        }
        break;

      case "missing_alt_text":
      case "missing_table_alt":
        var el = getElementById(slide, issue.elementId);
        if (el) {
          el.setTitle(issue.suggestedValue);
          el.setDescription(issue.suggestedValue);
        }
        break;

      case "empty_textbox":
        var el = getElementById(slide, issue.elementId);
        if (el) el.remove();
        break;

      // human-review only — mark as acknowledged
      case "small_text":
      case "no_speaker_notes":
      case "low_contrast":
        break;
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function remediateAll() {
  var result = scanPresentation();
  var fixed  = 0;
  var failed = 0;
  result.issues.forEach(function(issue) {
    if (!issue.autoFixable) return;
    var r = applyFix(issue);
    if (r.ok) fixed++; else failed++;
  });
  return { fixed: fixed, failed: failed };
}


// ── CLAUDE API ────────────────────────────────────────────────────────────────

function callClaude(prompt, imageBase64, mediaType) {
  var apiKey = getApiKey();
  if (!apiKey) return null;

  var content = [];
  if (imageBase64) {
    content.push({
      type: "image",
      source: {
        type:       "base64",
        media_type: mediaType || "image/png",
        data:       imageBase64,
      },
    });
  }
  content.push({ type: "text", text: prompt });

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 150,
    messages:   [{ role: "user", content: content }],
  };

  try {
    var response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method:           "post",
      contentType:      "application/json",
      headers: {
        "x-api-key":         apiKey,
        "anthropic-version": "2023-06-01",
      },
      payload:           JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) return null;
    var data = JSON.parse(response.getContentText());
    return data.content[0].text.trim();
  } catch(e) {
    return null;
  }
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
  return key.slice(0, 10) + "…" + key.slice(-4);
}


// ── HELPERS ───────────────────────────────────────────────────────────────────

function newIssue(fields) {
  return {
    id:             Utilities.getUuid(),
    checkType:      fields.checkType      || "",
    slideIndex:     fields.slideIndex     !== undefined ? fields.slideIndex : -1,
    elementId:      fields.elementId      || "",
    severity:       fields.severity       || "warning",
    title:          fields.title          || "",
    description:    fields.description    || "",
    suggestedValue: fields.suggestedValue || "",
    autoFixable:    !!fields.autoFixable,
    status:         "pending",
  };
}

function getTitleShape(slide) {
  var els = slide.getPageElements();
  for (var i = 0; i < els.length; i++) {
    var el = els[i];
    if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
    var shape = el.asShape();
    var ph    = shape.getPlaceholderType();
    if (ph === SlidesApp.PlaceholderType.TITLE ||
        ph === SlidesApp.PlaceholderType.CENTERED_TITLE) {
      return shape;
    }
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
    var shape = el.asShape();
    var ph    = shape.getPlaceholderType();
    if (ph === SlidesApp.PlaceholderType.TITLE ||
        ph === SlidesApp.PlaceholderType.CENTERED_TITLE) return;
    var txt = shape.getText().asString().trim();
    if (txt) parts.push(txt);
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
  function ch(c) {
    c /= 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  }
  var lum = 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b);
  return lum > 0.55;
}

function toHex(rgb) {
  function pad(n) { return ("0" + Math.round(n).toString(16)).slice(-2); }
  return pad(rgb.getRed()) + pad(rgb.getGreen()) + pad(rgb.getBlue());
}
