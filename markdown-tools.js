// Written by Bret Jordan
// https://github.com/jordan2175/gdoc-markdown-tools
// Last updated 2019-09-27
// Version 1.0.2
// Apache 2.0 License


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
              .addItem('A.b.C.', 'addSimpleHeadingUpperDot')
             )
  .addItem('Clear Heading Numbers', 'clearHeadingNumbers')
  .addSeparator()
  .addItem('Table of Contents', 'addtoc')
  .addToUi();
}



function onInstall(e) {
  onOpen(e);
}

// #### 1.2.3
function addMarkdownHeadingNumbers(){
  headingNumbers(true, 1, false, true);
}

// ### 1.2.3.
function addMarkdownHeadingNumbersDot(){
  headingNumbers(true, 1, true, true);
}

// #### a.b.c
function addMarkdownHeadingLower(){
  headingNumbers(true, 2, false, true);
}

// ### a.b.c.
function addMarkdownHeadingLowerDot(){
  headingNumbers(true, 2, true, true);
}

// #### A.B.C
function addMarkdownHeadingUpper(){
  headingNumbers(true, 3, false, true);
}

// ### A.B.C.
function addMarkdownHeadingUpperDot(){
  headingNumbers(true, 3, true, true);
}


// 1.2.3
function addSimpleHeadingNumbers(){
  headingNumbers(true, 1, false, false);
}

// 1.2.3.
function addSimpleHeadingNumbersDot(){
  headingNumbers(true, 1, true, false);
}

// a.b.c
function addSimpleHeadingLower(){
  headingNumbers(true, 2, false, false);
}

// a.b.c.
function addSimpleHeadingLowerDot(){
  headingNumbers(true, 2, true, false);
}

// a.b.c
function addSimpleHeadingUpper(){
  headingNumbers(true, 3, false, false);
}

// a.b.c.
function addSimpleHeadingUpperDot(){
  headingNumbers(true, 3, true, false);
}


// Clear all heading numbers
function clearHeadingNumbers(){
  headingNumbers(false, 0, false, false);
}







// Header Types: 1=numbers, 2=lowercase letters, 3=uppercase letters
function headingNumbers(add, headertype, enddot, markdown){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var paragraphs = document.getParagraphs();
  
  // This array of numbers will increment as we encounter new headings and thus keep 
  // the state of where we are at.  This is how the code knows which number is next.
  var numbers = [0,0,0,0,0,0,0];
  
  // Loop through all of the paragraphs that we have found
  for (var i in paragraphs) {
    var element = paragraphs[i];
    var text = element.getText()+'';
    var type = element.getHeading()+'';
    
    // exclude everything but headings
    if (!type.match(/Heading \d/)) {
      continue;
    }
    
    // exclude empty headings (e.g. page breaks generate these)
    if( text.match(/^\s*$/)){
      continue;
    }

    if (add == true) {
      // Get the current header level (e.g., 1, 2, 3, 4, 5, 6)
      var headerlevel = new RegExp(/Heading (\d)/).exec(type)[1];
      
      // This string will hold the actual header numbering value, be it integers or letters
      var numbering = '';
      
      // Increment the heading number in the array. This will enable us to know 
      // where we are at in the numbering.
      numbers[headerlevel]++;
      
      // Loop through all 6 possible levels and concatinate them in to a single string
      // Check for the type level and as needed, convert the numbers array in to strings
      // either numerical strings or letters.
      for (var level = 1; level <= 6; level++) {
        if (level < headerlevel) {
          if (headertype == 1) {
            numbering += numbers[level] + '.';
          } else if (headertype == 2) {
            numbering += int2lower(numbers[level]) + ".";
          } else if (headertype == 3) {
            numbering += int2upper(numbers[level]) + ".";
          }
        } else if (level == headerlevel) {
          if (enddot) {
            if (headertype == 1) {
              numbering += numbers[level] + '.';
            } else if (headertype == 2) {
              numbering += int2lower(numbers[level]) + ".";
            } else if (headertype == 3) {
              numbering += int2upper(numbers[level]) + ".";
            }
          } else {
            if (headertype == 1) {
              numbering += numbers[level];
            } else if (headertype == 2) {
              numbering += int2lower(numbers[level]);
            } else if (headertype == 3) {
              numbering += int2upper(numbers[level]);
            }
          }
        } else {
          numbers[level] = 0;
        }
      }

      // If there is an exsiting header number with out without a markdown hash, remove it.
      // The first part of the regex looks for an existing markdown header hash mark
      // Then it checks to see if there is an existing header number
      element.replaceText('^(#+\\s)?([0-9a-zA-Z\\-])(\\.[0-9a-zA-Z\\-])*\\.?\\s', '');
      
      // If this is for markdown, then we need to add the hash prefix for each heading level.
      var mdhash = ["", "","","","","",""];
      if (markdown == true) {
        mdhash = ["", "# ","## ","### ","#### ","##### ","###### "];
      }
 
      // Check each heading level and set the right numbers and apply the right style 
      if (headerlevel == 1) {
        element.insertText(0, mdhash[1] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      } else if (headerlevel == 2) {
        element.insertText(0, mdhash[2] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      } else if (headerlevel == 3) {
        element.insertText(0, mdhash[3] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      } else if (headerlevel == 4) {
        element.insertText(0, mdhash[4] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      } else if (headerlevel == 5) {
        element.insertText(0, mdhash[5] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING5);
      } else if (headerlevel == 6) {
        element.insertText(0, mdhash[6] + numbering + ' ');
        element.setHeading(DocumentApp.ParagraphHeading.HEADING6);
      }
    } else {
      // Clear current header numbers
      // If there is an exsiting header number with out without a markdown hash, remove it.
      // The first part of the regex looks for an existing markdown header hash mark
      // Then it checks to see if there is an existing header number
      element.replaceText('^(#+\\s)?([0-9a-zA-Z\\-])(\\.[0-9a-zA-Z\\-])*\\.?\\s', '');
    }
  }
}





function addtoc(){
  var document = DocumentApp.getActiveDocument();
  var body = document.getBody();
  var cursor = document.getCursor();
  var paragraphs = document.getParagraphs();
  
  h1data = body.getHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1);
  
  // The markdown toc indentation and prefix elements for each of the 6 heading levels.
  var tocindentstyle = ["", "- ","  * ","    + ","      - ","        * ","          + "];
  
  tocheadertext = "# Table of Contents\n";
  var c1 = cursor.insertText(tocheadertext);
  //c1.setFontSize(h1data{'FONT_SIZE'});
  //c1.setBold(h1data{'bold'});
  //c1.setForegroundColor(h1data{'FOREGROUND_COLOR'});
  c1.setAttributes(h1data)
  var position = document.newPosition(c1, tocheadertext.length);
  document.setCursor(position);
  
  
  // Loop through all of the paragraphs that we have found
  for (var i in paragraphs) {
    var element = paragraphs[i];
    var text = element.getText()+'';
    var type = element.getHeading()+'';
    
    // exclude everything but headings
    if (!type.match(/Heading \d/)) {
      continue;
    }
    
    // exclude empty headings (e.g. page breaks generate these)
    if( text.match(/^\s*$/)){
      continue;
    }

    // Get the current header level (e.g., 1, 2, 3, 4, 5, 6)
    var headerlevel = new RegExp(/Heading (\d)/).exec(type)[1];

    // We need to remove the hashes from our TOC text and TOC links 
    text = text.replace(/^#*\s?/, '');

    // We need to clean up the TOC link text a bit
    textlink = text.toLowerCase();
    textlink = textlink.replace(/([0-9a-zA-Z\-]\.?)+\s/, '');
    textlink = textlink.replace(/\s+/g, '-');
   
    if (cursor) {
      // Check each heading level and set the right indentation for the toc 
      if (headerlevel == 1) {
        toctextline = tocindentstyle[1] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        t.setFontSize(10);
        t.setForegroundColor("#333333");
        t.setBold(false);
        t.setFontFamily("Consolas");
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      } else if (headerlevel == 2) {
        toctextline = tocindentstyle[2] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      } else if (headerlevel == 3) {
        toctextline = tocindentstyle[3] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      } else if (headerlevel == 4) {
        toctextline = tocindentstyle[4] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      } else if (headerlevel == 5) {
        toctextline = tocindentstyle[5] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      } else if (headerlevel == 6) {
        toctextline = tocindentstyle[6] + '[' + text + '](#' + textlink + ')\n';
        var t = cursor.insertText(toctextline);
        var position = document.newPosition(t, toctextline.length);
        document.setCursor(position);
      }
    }
  }
}



// --------------------------------------------------------------------------------
// Supporting functions 
// --------------------------------------------------------------------------------

// Convert integer to lower case letter
function int2lower(num){
  if ( (num < 1) || (num > 27) ) { num = 45 - 96; }
  return String.fromCharCode(96+num); 
}

// Convert integer to upper case letter
function int2upper(num){
  if ( (num < 1) || (num > 27) ) { num = 45 - 64; }
  return String.fromCharCode(64+num); 
}

