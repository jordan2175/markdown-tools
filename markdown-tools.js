// Written by Bret Jordan
// https://github.com/jordan2175/markdown-tools
// Last updated 2025-05-27
// Version 2.0 (V8 runtime compatible)
// Apache 2.0 License
 
// onOpen 
function onOpen(e) {
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
  .addSubMenu(ui.createMenu('Heading Numbers (Markdown)')
              .addItem('1.2.3', 'addMarkdownHeadingNumbers')
              .addItem('1.2.3.', 'addMarkdownHeadingNumbersDot')
              .addItem('a.b.c', 'addMarkdownHeadingLower')
              .addItem('a.b.c.', 'addMarkdownHeadingLowerDot')
              .addItem('A.B.C', 'addMarkdownHeadingUpper')
              .addItem('A.B.C.', 'addMarkdownHeadingUpperDot')
             )
  .addSubMenu(ui.createMenu('Heading Numbers')
              .addItem('1.2.3', 'addSimpleHeadingNumbers')
              .addItem('1.2.3.', 'addSimpleHeadingNumbersDot')
              .addItem('a.b.c', 'addSimpleHeadingLower')
              .addItem('a.b.c.', 'addSimpleHeadingLowerDot')
              .addItem('A.B.C', 'addSimpleHeadingUpper')
              .addItem('A.B.C.', 'addSimpleHeadingUpperDot')
             )
  .addItem('Clear Heading Numbers', 'clearHeadingNumbers')
  .addSeparator()
  .addItem('Table of Contents', 'addtoc')
  .addSeparator()
  .addItem('About', 'showAbout')
  .addToUi();
}
 
 
// ----------------------------------------
// Run on install
// ----------------------------------------
 
function onInstall(e) {
  onOpen(e);
}
 
 
// ----------------------------------------
// FIX: In V8, element.getHeading() returns an enum object, not a string.
// Concatenating with '' now yields e.g. "HEADING1" instead of "Heading 1",
// so the original /Heading \d/ regex never matched anything.
// Solution: build a map from each enum value to a numeric level (1-6),
// then compare with === instead of regex string matching.
// ----------------------------------------
 
function buildHeadingLevelMap() {
  var ph = DocumentApp.ParagraphHeading;
  var map = new Map();
  map.set(ph.HEADING1, 1);
  map.set(ph.HEADING2, 2);
  map.set(ph.HEADING3, 3);
  map.set(ph.HEADING4, 4);
  map.set(ph.HEADING5, 5);
  map.set(ph.HEADING6, 6);
  return map;
}
 
 
// ----------------------------------------
// Menu Functions
// ----------------------------------------
 
function addMarkdownHeadingNumbers()    { headingNumbers(true,  1, false, true);  }
function addMarkdownHeadingNumbersDot() { headingNumbers(true,  1, true,  true);  }
function addMarkdownHeadingLower()      { headingNumbers(true,  2, false, true);  }
function addMarkdownHeadingLowerDot()   { headingNumbers(true,  2, true,  true);  }
function addMarkdownHeadingUpper()      { headingNumbers(true,  3, false, true);  }
function addMarkdownHeadingUpperDot()   { headingNumbers(true,  3, true,  true);  }
 
function addSimpleHeadingNumbers()      { headingNumbers(true,  1, false, false); }
function addSimpleHeadingNumbersDot()   { headingNumbers(true,  1, true,  false); }
function addSimpleHeadingLower()        { headingNumbers(true,  2, false, false); }
function addSimpleHeadingLowerDot()     { headingNumbers(true,  2, true,  false); }
function addSimpleHeadingUpper()        { headingNumbers(true,  3, false, false); }
function addSimpleHeadingUpperDot()     { headingNumbers(true,  3, true,  false); }
 
function clearHeadingNumbers()          { headingNumbers(false, 0, false, false); }
 
 
function showAbout() {
  var template = HtmlService.createTemplateFromFile("AboutDialog");
  var page = template
              .evaluate()
              .setSandboxMode(HtmlService.SandboxMode.IFRAME)
              .setWidth(500)
              .setHeight(210);
  DocumentApp.getUi()
      .showModalDialog(page, "About Markdown Tools");
}
 
 
// ----------------------------------------
// Main Functions
// ----------------------------------------
 
// Header Types: 1=numbers, 2=lowercase letters, 3=uppercase letters
function headingNumbers(add, headertype, enddot, markdown) {
  var document   = DocumentApp.getActiveDocument();
  var body       = document.getBody();
  var paragraphs = body.getParagraphs();  // FIX: was document.getParagraphs()
  var levelMap   = buildHeadingLevelMap(); // FIX: replaces string-based heading detection
 
  var numbers = [0, 0, 0, 0, 0, 0, 0];
 
  for (var i = 0; i < paragraphs.length; i++) {
    var element     = paragraphs[i];
    var text        = element.getText();
    var headingEnum = element.getHeading(); // returns enum, not string
 
    // FIX: compare enum value directly instead of regex on string representation
    var headerlevel = levelMap.get(headingEnum);
    if (!headerlevel) continue; // not a heading
 
    // exclude empty headings (e.g. page breaks generate these)
    if (text.match(/^\s*$/)) continue;
 
    if (add) {
      var numbering = '';
      numbers[headerlevel]++;
 
      for (var level = 1; level <= 6; level++) {
        if (level < headerlevel) {
          if      (headertype === 1) { numbering += numbers[level] + '.'; }
          else if (headertype === 2) { numbering += int2lower(numbers[level]) + '.'; }
          else if (headertype === 3) { numbering += int2upper(numbers[level]) + '.'; }
        } else if (level === headerlevel) {
          if (enddot) {
            if      (headertype === 1) { numbering += numbers[level] + '.'; }
            else if (headertype === 2) { numbering += int2lower(numbers[level]) + '.'; }
            else if (headertype === 3) { numbering += int2upper(numbers[level]) + '.'; }
          } else {
            if      (headertype === 1) { numbering += numbers[level]; }
            else if (headertype === 2) { numbering += int2lower(numbers[level]); }
            else if (headertype === 3) { numbering += int2upper(numbers[level]); }
          }
        } else {
          numbers[level] = 0;
        }
      }
 
      // FIX: replaceText() in V8 uses the RE2 engine (not Java regex).
      // The original character class syntax [(0-9)*\-] was malformed and
      // relied on Java regex leniency. Rewritten as clean RE2-compatible patterns.
 
      // Remove existing heading number (with or without leading markdown hashes)
      element.replaceText('^(#+ )?[0-9]+([.][0-9]+)*[.]? ', '');
      // Remove existing heading letter (with or without leading markdown hashes)
      element.replaceText('^(#+ )?[a-zA-Z]([.][a-zA-Z])*[.]? ', '');
 
      var mdhash = ["", "", "", "", "", "", ""];
      if (markdown) {
        mdhash = ["", "# ", "## ", "### ", "#### ", "##### ", "###### "];
      }
 
      var headingConst = [
        null,
        DocumentApp.ParagraphHeading.HEADING1,
        DocumentApp.ParagraphHeading.HEADING2,
        DocumentApp.ParagraphHeading.HEADING3,
        DocumentApp.ParagraphHeading.HEADING4,
        DocumentApp.ParagraphHeading.HEADING5,
        DocumentApp.ParagraphHeading.HEADING6
      ];
 
      element.insertText(0, mdhash[headerlevel] + numbering + ' ');
      element.setHeading(headingConst[headerlevel]);
 
    } else {
      // Clear mode: remove existing heading numbers/letters
      element.replaceText('^(#+ )?[0-9]+([.][0-9]+)*[.]? ', '');
      element.replaceText('^(#+ )?[a-zA-Z]([.][a-zA-Z])*[.]? ', '');
    }
  }
}
 
 
function addtoc() {
  var document  = DocumentApp.getActiveDocument();
  var selection = document.getSelection();
  if (selection) return; // TOC goes at cursor, not selection
 
  var body       = document.getBody();
  var cursor     = document.getCursor();
  var paragraphs = body.getParagraphs(); // FIX: was document.getParagraphs()
  var levelMap   = buildHeadingLevelMap(); // FIX: replaces string-based heading detection
 
  var h1data = body.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1);
 
  var tocindentstyle = ["", "- ", "  * ", "    + ", "      - ", "        * ", "          + "];
 
  // FIX: cursor.insertText() walk no longer works reliably in V8.
  // Find the body-child index of the cursor's paragraph, then use
  // body.insertParagraph(index, text) to insert each TOC line in order.
  var cursorElement = cursor.getElement();
  while (cursorElement.getParent() &&
         cursorElement.getParent().getType() !== DocumentApp.ElementType.BODY_SECTION) {
    cursorElement = cursorElement.getParent();
  }
  var insertIndex = body.getChildIndex(cursorElement);
 
  // Insert TOC header
  var tocHeader = body.insertParagraph(insertIndex, '# Table of Contents');
  tocHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  tocHeader.setAttributes(h1data);
  insertIndex++;
 
  for (var i = 0; i < paragraphs.length; i++) {
    var element     = paragraphs[i];
    var text        = element.getText();
    var headingEnum = element.getHeading();
 
    // FIX: use enum map instead of string regex
    var headerlevel = levelMap.get(headingEnum);
    if (!headerlevel) continue;
    if (text.match(/^\s*$/)) continue;
 
    // Strip leading markdown hashes from display text
    text = text.replace(/^#+ /, '');
 
    // Build anchor: strip heading numbers/letters, lowercase, hyphens for spaces
    var textlink = text.toLowerCase();
    textlink = textlink.replace(/^[0-9a-zA-Z]([.][0-9a-zA-Z])*[.]? /, '');
    textlink = textlink.replace(/\s+/g, '-');
 
    var toctextline = tocindentstyle[headerlevel] + '[' + text + '](#' + textlink + ')';
 
    var tocPara = body.insertParagraph(insertIndex, toctextline);
    tocPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    tocPara.setFontSize(10);
    tocPara.setForegroundColor("#333333");
    tocPara.setBold(false);
    tocPara.setFontFamily("Consolas");
 
    // Link just the visible text between [ and ]
    var linkStart = tocindentstyle[headerlevel].length + 1;
    var linkEnd   = linkStart + text.length - 1;
    tocPara.editAsText().setLinkUrl(linkStart, linkEnd, '#' + textlink);
 
    insertIndex++;
  }
}
 
 
// --------------------------------------------------------------------------------
// Supporting functions
// --------------------------------------------------------------------------------
 
function int2lower(num) {
  if ((num < 1) || (num > 26)) { num = 45 - 96; }
  return String.fromCharCode(96 + num);
}
 
function int2upper(num) {
  if ((num < 1) || (num > 26)) { num = 45 - 64; }
  return String.fromCharCode(64 + num);
}
