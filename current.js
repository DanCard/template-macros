function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Custom Macros')
    .addItem('Format and set Title', 'formatAndSetTitle')
    .addItem('Format links', 'formatLinks')
    .addItem('Format code', 'formatCode')
    .addItem('Format table at cursor', 'formatCurrentTable')
/*
    .addItem('Set title', 'setTitle')
    .addItem('Format title font size 30', 'formatTitle')
*/
    .addToUi();

  var doc = DocumentApp.getActiveDocument();
  const name = doc.getName();

  // Don't execute setting a title if title already exists.
  if (name == "template macros" ||
      name.indexOf('-' ) > -1   ||
      name.length > ('untitled document').length ||
      name.indexOf('  ') > -1) return;

  // For changes or debugging, need to save first with floppy disk icon above.

  // Get the current date
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), '  -  EEE d MMM yy  h:mm aaa');
  doc.setName(dateStr);
  body = doc.getBody();
  paragraph = body.insertParagraph(0, dateStr);
  // Set the document title to the current date
  paragraph.setHeading(DocumentApp.ParagraphHeading.TITLE);
}

function formatAndSetTitle() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const firstParagraph = body.getParagraphs()[0];
  
  if (!firstParagraph || firstParagraph.getText().trim() === "") {
    DocumentApp.getUi().alert('The first line of the document appears to be empty.');
    return;
  }

  const textElement = firstParagraph.editAsText();
  const textString = textElement.getText();
  
  // 1. Identify the Delimiter and Split Point
  const delimiterRegex = /[?!-]/g;
  let match, lastMatch;
  while ((match = delimiterRegex.exec(textString)) !== null) {
    lastMatch = match;
  }

  let splitIndex = -1;
  let titleEndIndex = -1;

  if (lastMatch) {
    splitIndex = lastMatch.index;
    // For dashes, title ends before the char. For ? and !, title includes the char.
    titleEndIndex = (lastMatch[0] === '-') ? splitIndex - 1 : splitIndex;
  } else {
    // Fallback: Search for the date pattern (e.g., "Sat 31 Jan 26")
    const dateRegex = /\s([A-Z][a-z]{2}\s\d{1,2}\s[A-Z][a-z]{2}\s\d{2})/;
    const dateMatch = textString.match(dateRegex);

    if (dateMatch) {
      splitIndex = dateMatch.index;
      titleEndIndex = splitIndex - 1;
    } else {
      DocumentApp.getUi().alert('No title delimiter (?, !, -) or date pattern found.');
      return;
    }
  }

  const titlePart = textString.substring(0, titleEndIndex + 1);
  const datePart = textString.substring(titleEndIndex + 1);

  // --- CONFIGURATION ---
  const MAX_TITLE_SIZE = 34;
  const PAGE_WIDTH_POINTS = 495; 
  const CHAR_WIDTH_FACTOR = 0.45; 
  const MIN_FONT_SIZE = 6;

  // Approximate width calculation
  const getEstimatedWidth = (tSize, dSize) => {
    return (titlePart.length * tSize * CHAR_WIDTH_FACTOR) + 
           (datePart.length * dSize * CHAR_WIDTH_FACTOR);
  };

  let currentFontSize = textElement.getFontSize(0) || 12;
  let titleFontSize = currentFontSize;
  let dateFontSize = 12;

  // STEP 1: Shrink Logic
  // If the line currently wraps, shrink both parts until it fits
  if (getEstimatedWidth(titleFontSize, dateFontSize) > PAGE_WIDTH_POINTS) {
    while (getEstimatedWidth(titleFontSize, dateFontSize) > PAGE_WIDTH_POINTS && titleFontSize > MIN_FONT_SIZE) {
      titleFontSize--;
      dateFontSize--;
    }
  } 
  // STEP 2: Grow Logic
  // If it fits, increase ONLY the Title size (up to 30)
  else {
    let tempTitleSize = titleFontSize;
    while (getEstimatedWidth(tempTitleSize + 1, dateFontSize) <= PAGE_WIDTH_POINTS && tempTitleSize < MAX_TITLE_SIZE) {
      tempTitleSize++;
    }
    titleFontSize = tempTitleSize;
  }

  // APPLY FORMATTING
  // Clear any existing size formatting first to avoid overlaps
  textElement.setFontSize(0, textString.length - 1, dateFontSize);
  
  // Apply the specific title size
  textElement.setFontSize(0, titleEndIndex, titleFontSize);
  
  // Ensure the date part is the correct size (already set by the clear line above, but for clarity:)
  textElement.setFontSize(titleEndIndex + 1, textString.length - 1, dateFontSize);

  setTitle();
}

function resizeTitle2() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const firstParagraph = body.getParagraphs()[0];
  
  if (!firstParagraph || firstParagraph.getText().trim() === "") {
    DocumentApp.getUi().alert('The first line of the document appears to be empty.');
    return;
  }

  const textElement = firstParagraph.editAsText();
  const textString = textElement.getText();
  
  // Split the text into Title and Date using " - " as the delimiter
  const splitIndex = textString.lastIndexOf(" - ");
  
  // If the format doesn't match, we can't reliably split it.
  if (splitIndex === -1) {
    DocumentApp.getUi().alert('Format not found. Ensure the line ends with " - <date>".');
    return;
  }

  const titlePart = textString.substring(0, splitIndex);
  const datePart = textString.substring(splitIndex);

  // --- CONFIGURATION ---
  const MAX_TITLE_SIZE = 30;
  const PAGE_WIDTH_POINTS = 502; // by trial and error
  const CHAR_WIDTH_FACTOR = 0.44; // Tuned for Spectral Light (thin serif)
  const MIN_FONT_SIZE = 9;

  // Approximate width calculation
  const getEstimatedWidth = (tSize, dSize) => {
    return (titlePart.length * tSize * CHAR_WIDTH_FACTOR) + 
           (datePart.length * dSize * CHAR_WIDTH_FACTOR);
  };

  let currentFontSize = textElement.getFontSize(0) || 12;
  let titleFontSize = currentFontSize;
  let dateFontSize = 12;

  // STEP 1: Shrink Logic
  // If it currently doesn't fit, shrink both Title and Date equally
  if (getEstimatedWidth(titleFontSize, dateFontSize) > PAGE_WIDTH_POINTS) {
    while (getEstimatedWidth(titleFontSize, dateFontSize) > PAGE_WIDTH_POINTS && titleFontSize > MIN_FONT_SIZE) {
      titleFontSize--;
      dateFontSize--;
    }
  }

  // STEP 2: Grow Logic.   If it fits, increase ONLY the Title size up to 30
  else {
    let tempTitleSize = titleFontSize;
    while (getEstimatedWidth(tempTitleSize + 1, dateFontSize) <= PAGE_WIDTH_POINTS && tempTitleSize < MAX_TITLE_SIZE) {
      tempTitleSize++;
    }
    titleFontSize = tempTitleSize;
  }

  // APPLY FORMATTING
  // 1. Title Size
  textElement.setFontSize(0, splitIndex - 1, titleFontSize);
  // 2. Date Size (everything from the " - " onwards)
  textElement.setFontSize(splitIndex, textString.length - 1, dateFontSize);

  setTitle();
}


function setTitle() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  paragraph = body.getParagraphs()[0];
  doc.setName(paragraph.getText());
}

function formatLinks() {
  var body = DocumentApp.getActiveDocument().getBody();

  // Basic cleanup
  body.replaceText('https://www\.', '');
  body.replaceText('\.com/', '/');
  body.replaceText('facebook/', '');
  body.replaceText('https://', '');

  var paragraphs = body.getParagraphs();
  paragraphs.forEach(function(paragraph) {
    const textElement = paragraph.editAsText();
    var text = textElement.getText();
    var isUnderlined = false;

    // Check if any of the text is underlined (common for links)
    for (let j = 0; j < text.length; j++) {
      if (textElement.isUnderline(j)) {
        isUnderlined = true;
        break;
      }
    }

    if (isUnderlined) {
      // Replace slashes with spaces and remove underline
      textElement.replaceText('/', ' ');
      textElement.setUnderline(false);
      text = textElement.getText(); // Refresh text after changes
    }

    // --- NEW: Handle lmarena links ---
    // Target: lmarena.ai -> lm arena ai
    if (text.indexOf('lmarena.ai') !== -1) {
      // 1. Replace "lmarena.ai" with "lm arena ai" (adds space between lm/arena and arena/ai)
      textElement.replaceText('lmarena\\.ai', 'lm arena ai');
      text = textElement.getText(); // Refresh text for index calculation

      var lmRegex = /lm arena ai/g;
      var lmMatch;
      while ((lmMatch = lmRegex.exec(text)) !== null) {
        var start = lmMatch.index;
        
        // "lm " = 3 chars. "arena" = 5 chars.
        // Format "arena" to 30pt (Indices: start + 3 to start + 7)
        textElement.setFontSize(start + 3, start + 7, 30);

        // "lm arena ai" = 11 chars. 
        // Format everything AFTER "ai" to 8pt
        var afterAiIndex = start + 11;
        if (afterAiIndex < text.length) {
          textElement.setFontSize(afterAiIndex, text.length - 1, 8);
        }
      }
    }

    if ((startIndex = text.indexOf('- Wikipedia')) !== -1) {
      //                            12345678901
      textElement.setFontSize(startIndex, startIndex + 10, 15);
    }

    // --- Existing: Handle aistudio links ---
    if (text.indexOf('aistudio.google') !== -1) {
      textElement.replaceText('aistudio\.google', 'ai studio google');
      text = textElement.getText();

      var aiRegex = /ai studio google/g;
      var match;
      while ((match = aiRegex.exec(text)) !== null) {
        var startIndex = match.index;
        textElement.setFontSize(startIndex, startIndex + 1, 30);
        var afterGoogleIndex = startIndex + 16;
        if (afterGoogleIndex < text.length) {
          textElement.setFontSize(afterGoogleIndex, text.length - 1, 8);
        }
      }
    }

    // --- Existing: Format "status <number>" ---
    var statusRegex = /status (\d+)/g;
    var sMatch;
    while ((sMatch = statusRegex.exec(text)) !== null) {
      textElement.setFontSize(sMatch.index, sMatch.index + sMatch[0].length - 1, 8);
    }
  });
}

function formatCode() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var elementType = DocumentApp.ElementType;
  if (selection) {
    var elements = selection.getRangeElements();
    elements.forEach(function(element) {
      var type = element.getElement().getType();
      // if (type === DocumentApp.ElementType.TEXT) {
      //   element = element.getParent();
      //   console.log('element : ' + element)
      //   type = element.getElement().getType();
      //   console.log('type : ' + type)
      // }
      if (type == elementType.PARAGRAPH) {
        var paragraph = element.getElement().asParagraph();
        paragraph.setFontFamily("Spectral;300");  // 300 indicates Spectral Light, from trial and error
        paragraph.setFontSize(12);
        paragraph.setLineSpacing(0.7);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);      
      }
    });
  } else {
    // DocumentApp.getUi().alert('Macro works on a selection.  Nothing selected.');
    // Assume misclick and I want to format links.
    formatLinks();
  }
}

function setBeforeAfterSpacingOne() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var elementType = DocumentApp.ElementType;
  if (selection) {
    var elements = selection.getRangeElements();
    elements.forEach(function(element) {
      var type = element.getElement().getType();
      // if (type === DocumentApp.ElementType.TEXT) {
      //   element = element.getParent();
      //   console.log('element : ' + element)
      //   type = element.getElement().getType();
      //   console.log('type : ' + type)
      // }
      if (type == elementType.PARAGRAPH) {
        var paragraph = element.getElement().asParagraph();
        paragraph.setSpacingAfter(1);
        paragraph.setSpacingBefore(1);      
      }
    });
  } else {
    // DocumentApp.getUi().alert('Macro works on a selection.  Nothing selected.');
    // Assume misclick and I want to format links.
    formatLinks();
  }
}


/**
 * Takes the user's selected text and sets the line spacing to 0.7,
 * and the paragraph spacing before and after to 9 points.
 *
 * This version uses the most aggressive method: it first clears all
 * formatting from the element and then applies the new attributes.
 */
function formatSelection() {
  const selection = DocumentApp.getActiveDocument().getSelection();

  if (selection) {
    // Define the final, correct formatting attributes.
    const finalAttributes = {};
    finalAttributes[DocumentApp.Attribute.LINE_SPACING] = 0.7;
    finalAttributes[DocumentApp.Attribute.SPACING_BEFORE] = 9;
    finalAttributes[DocumentApp.Attribute.SPACING_AFTER] = 10;

    // We need to collect the unique paragraphs/list items to format.
    // A selection can span multiple text runs in the same paragraph.
    const elementsToFormat = {};

    const rangeElements = selection.getRangeElements();
    for (const rangeElement of rangeElements) {
      let container = rangeElement.getElement();
      
      // If the element is Text, we need to format its parent.
      if (container.getType() === DocumentApp.ElementType.TEXT) {
        container = container.getParent();
      }

      // We only care about elements that can have paragraph-level attributes.
      if (container.setAttributes) {
        // Use the element's unique ID to avoid formatting the same one multiple times.
        const elementId = container.getAttributes()['ELEMENT_ID']; // A trick to get a unique ID
        if (!elementId) {
           // Fallback for elements without a stable ID
           elementsToFormat[Math.random()] = container;
        } else {
           elementsToFormat[elementId] = container;
        }
      }
    }

    // Now, iterate over the unique elements and format them.
    for (const id in elementsToFormat) {
      const element = elementsToFormat[id];
      try {
        // STEP 1: THE NUCLEAR OPTION. Clear all formatting.
        // An empty attributes object tells the API to remove all styling.
        element.setAttributes({});

        // STEP 2: REBUILD. Apply the correct formatting to the clean slate.
        element.setAttributes(finalAttributes);

      } catch (e) {
        console.error(`Could not format element of type ${element.getType()}: ${e.message}`);
        DocumentApp.getUi().alert(`An error occurred: ${e.message}`);
      }
    }
    
  } else {
    DocumentApp.getUi().alert('Please select some text first.');
  }
}


/**
 * Takes the user's selected text and sets the line spacing to 0.7,
 * and the paragraph spacing before and after to 9 points.
 *
 * This version includes a specific workaround for a known bug where
 * formatting does not apply to new ListItem elements.
 */
function formatSelectionVerticalSpacing() {
  const selection = DocumentApp.getActiveDocument().getSelection();

  if (selection) {
    // Define the final, correct formatting attributes.
    const finalAttributes = {};
    finalAttributes[DocumentApp.Attribute.LINE_SPACING] = 0.7;
    finalAttributes[DocumentApp.Attribute.SPACING_BEFORE] = 10;
    finalAttributes[DocumentApp.Attribute.SPACING_AFTER] = 10;

    const rangeElements = selection.getRangeElements();

    for (const rangeElement of rangeElements) {
      // We only care about the element if it's not a partial selection.
      // If it is partial, getElement() will be the Text, and we'll get its parent.
      if (!rangeElement.isPartial()) {
        const element = rangeElement.getElement();
        
        // Check if the element is a Paragraph or ListItem that can be formatted.
        if (element.setAttributes) {
          applyFormatting(element, finalAttributes);
        }
      } else {
        // For partial selections, get the parent container (Paragraph or ListItem).
        const container = rangeElement.getElement().getParent();
        if (container.setAttributes) {
          applyFormatting(container, finalAttributes);
        }
      }
    }
  } else {
    DocumentApp.getUi().alert('Please select some text first.');
  }
  setSpacingBeforeAfter()
}

/**
 * Helper function to apply formatting attributes, containing the bug workaround.
 * @param {GoogleAppsScript.Document.Paragraph|GoogleAppsScript.Document.ListItem} element The element to format.
 * @param {Object} attributes The attributes object to apply.
 */
function applyFormatting(element, attributes) {
  try {
    // === THE BUG WORKAROUND ===
    // For ListItem elements, the API often fails to apply attributes directly
    // because the item inherits styles from the parent List.
    // By setting a single attribute first (even to a temporary value),
    // we "detach" the item from the default style, forcing it to accept
    // the new set of attributes immediately after.
    if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
      // The "nudge": set a value to force the element to have its own attributes.
      element.setSpacingBefore(0); 
    }
    
    // Now, apply the final, correct attributes. This will now work.
    element.setAttributes(attributes);

  } catch (e) {
    console.error(`Could not format element of type ${element.getType()}: ${e.message}`);
    DocumentApp.getUi().alert(`An error occurred: ${e.message}`);
  }
}

/**
 * Takes the user's selected text and sets the line spacing to 0.7,
 * and the paragraph spacing before and after to 9 points.
 * This version uses setAttributes() for greater reliability, especially with list items.
 */
function setSpacingBeforeAfter() {
  setParagraphSpacing()
  // Get the active document's user selection.
  const selection = DocumentApp.getActiveDocument().getSelection();

  if (selection) {
    // Define the formatting attributes we want to apply.
    // Using the Attribute enum is best practice.
    const formattingAttributes = {};
    formattingAttributes[DocumentApp.Attribute.LINE_SPACING] = 0.7;
    formattingAttributes[DocumentApp.Attribute.SPACING_BEFORE] = 8;
    formattingAttributes[DocumentApp.Attribute.SPACING_AFTER] = 8;

    // Get the individual elements within the selection.
    const rangeElements = selection.getRangeElements();

    for (const rangeElement of rangeElements) {
      const element = rangeElement.getElement();
      let container = element;

      // If the selected element is Text, get its parent (Paragraph or ListItem).
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        container = element.getParent();
      }

      // Check if the container element is a type that supports attributes.
      // This will work for both PARAGRAPH and LIST_ITEM.
      if (container.setAttributes) {
        try {
          // Apply all attributes at once. This is more robust.
          container.setAttributes(formattingAttributes);
        } catch (e) {
          console.error(`Could not format element of type ${container.getType()}: ${e.message}`);
        }
      }
    }
  } else {
    // If nothing is selected, inform the user.
    DocumentApp.getUi().alert('Please select some text first.');
  }
}

/**
 * Takes the user's selected text and sets the line spacing to 0.7,
 * and the paragraph spacing before and after to 9 points.
 * 
 * Similar to function formatCode()
 */
function setParagraphSpacing() {
  // Get the active document's user selection.
  const selection = DocumentApp.getActiveDocument().getSelection();

  if (selection) {
    // Get the individual elements within the selection (like text, paragraphs, etc.).
    const rangeElements = selection.getRangeElements();

    for (const rangeElement of rangeElements) {
      const element = rangeElement.getElement();
      console.log('element: ' + element)
      let container = element;

      // If the selected element is a Text element (i.e., a partial selection),
      // we need to get its parent paragraph or list item to apply formatting.
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        container = element.getParent();
        console.log('element.getParent();  container: ' + container);
      }

      // Check if the container element is a type that supports paragraph-level
      // formatting (like PARAGRAPH or LIST_ITEM).
      // This is a robust way to ensure we don't try to format something unsupported.
      if (container.setLineSpacing && container.setSpacingBefore && container.setSpacingAfter) {
        console.log('try set...Spacing...()')
        try {
          // Apply the desired formatting.
          container.setLineSpacing(0.7);
          console.log('Done with: container.setLineSpacing(0.7);')
        } catch (e) {
          console.error(`Could not setLineSpacing(0.7) of type ${container.getType()}: ${e.message}`);
        }
        try {
          container.setSpacingBefore(9);
          console.log('Done with: container.setSpacingBefore();')
        } catch (e) {
          console.error(`Could not setSpacingBefore() for element of type ${container.getType()}: ${e.message}`);
        }
        try {
          container.setSpacingAfter(10);
          console.log('Done with: container.setSpacingAfter();')
        } catch (e) {
          console.error(`Could not setSpacingAfter() for element of type ${container.getType()}: ${e.message}`);
        }
      }
    }
  } else {
    // If nothing is selected, inform the user.
    DocumentApp.getUi().alert('Please select some text first.');
  }
}



/**
 * Formats the table the cursor is currently in.
 *
 * 1. Removes the minimum row height option.
 * 2. Decreases the font size of the entire table by 2 points.
 * 3. Sets line and paragraph spacing to 2 points before and after for text in each cell.
 * 4. Sets cell padding to 0.02 inches for all cells.
 */
function formatCurrentTable() {
  try {
    // 1. Get the current document and cursor position.
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();

    if (!cursor) {
      Logger.log("No cursor found in the document.");
      DocumentApp.getUi().alert("Error", "No cursor found in the document.", DocumentApp.getUi().ButtonSet.OK);
      return;
    }

    // 2. Find the table containing the cursor. Iterate upwards through parent elements.
    let element = cursor.getElement();
    let table = null;

    while (element) {
      if (element.getType() === DocumentApp.ElementType.TABLE) {
        table = element.asTable();
        break;  // Found the table, exit the loop.
      }
      element = element.getParent();
    }

    if (!table) {
      Logger.log("No table found at the cursor's position.");
      DocumentApp.getUi().alert("Error", "No table found at the cursor's position.", DocumentApp.getUi().ButtonSet.OK);
      return;
    }
    
    table.setBorderWidth(0.25)
    const headerRow = table.getRow(0);
    // Loop through each cell in the header row
    for (var col = 0; col < headerRow.getNumCells(); col++) {
      var cell = headerRow.getCell(col);

      // Set cell alignment to middle
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.MIDDLE);
      cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }

    // 3. Iterate through rows and cells to apply formatting

    const numRows = table.getNumRows();
    for (let i = 0; i < numRows; i++) {
      const row = table.getRow(i);
      const numCells = row.getNumCells();
      
      for (let j = 0; j < numCells; j++) {
        const cell = row.getCell(j);

        // Apply the formatting steps to the current cell:
        // -----------------------------------------------------
        
        // Step 1: Remove minimum row height (applied to the row containing the cell).
        row.setMinimumHeight(null);

        // Step 2: Decrease font size by 2 (to the text content of the cell).
        // Get the cell's text.
        const cellText = cell.getText();

        // Get the current font size (if any, otherwise default to 11)
        let currentFontSize = 16; // Default value.  Important to have a sensible default.
        const firstElement = cell.getChild(0);
        if (firstElement && firstElement.getType() === DocumentApp.ElementType.PARAGRAPH){
          const textStyle = firstElement.getAttributes();
          if(textStyle && textStyle[DocumentApp.Attribute.FONT_SIZE]){
             currentFontSize = textStyle[DocumentApp.Attribute.FONT_SIZE];
          }
        }

        // Ensure the font size does not go below 1pt.
        const newFontSize = Math.max(currentFontSize - 2, 1);

        // Apply the font size change to the entire cell text.
        cell.editAsText().setFontSize(newFontSize);

        // Step 3: Set line and paragraph spacing to 2 points before and after.
        // Iterate through all paragraphs in the cell.
        for (let k = 0; k < cell.getNumChildren(); k++) {
          const child = cell.getChild(k);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            paragraph.setSpacingBefore(2);
            paragraph.setSpacingAfter(2);
          }
        }

        // Step 4: Set cell padding
        cell.setPaddingTop(2);
        cell.setPaddingBottom(2);
        cell.setPaddingLeft(2);
        cell.setPaddingRight(2);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      }
    }

    // Refresh the UI (optional, but helps visualize the changes immediately).
    doc.saveAndClose();
    DocumentApp.openById(doc.getId()); // Reopen to refresh.
    
    Logger.log("Table formatting complete.");
  } catch (e) {
    Logger.log("Error formatting table: " + e.toString());
    DocumentApp.getUi().alert("Error", "Error formatting table: " + e.toString(), DocumentApp.getUi().ButtonSet.OK);
  }
}


function formatTitle() {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var paragraph = body.getParagraphs()[0];
    var text = paragraph.getText();
    var dashIndex = text.indexOf('-');

    if (dashIndex > -1) {
        // Format from start (0) to right before the dash (dashIndex)
        var textElement = paragraph.editAsText();
        textElement.setFontSize(0, dashIndex - 1, 30);
    }
}

