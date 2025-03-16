/**
 * Replaces variables in a Google Slides presentation with values from a form
 * and exports the presentation as a PDF.
 *
 * Variables are defined in the form ${variable_name}.
 */
function replaceVariablesAndExportPdfWithForm() {
  const presentation = SlidesApp.getActivePresentation();
  const presentationName = presentation.getName();
  const slides = presentation.getSlides();

  // Create a form to collect variable values
  const ui = SlidesApp.getUi();
  const formBuilder = ui.createForm('Enter Variable Values');

  // Extract variable names from the presentation
  const variableNames = extractVariableNames(slides);

  // Add text input fields to the form for each variable
  const variableValues = {};
  variableNames.forEach(name => {
    const response = ui.prompt(`Enter value for ${name}`);
    if (response.getSelectedButton() == ui.Button.OK) {
      variableValues[name] = response.getResponseText();
    } else {
      return; // Cancel the export if the user cancels the form
    }
  });

  // Replace variables in the presentation
  slides.forEach(slide => {
    const pageElements = slide.getPageElements();

    pageElements.forEach(element => {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        if (shape.getText()) {
          const textRange = shape.getText();
          let text = textRange.asString();

          for (const variableName in variableValues) {
            const variablePattern = `\\$\\{${variableName}\\}`;
            const regex = new RegExp(variablePattern, 'g');
            text = text.replace(regex, variableValues[variableName]);
          }

          textRange.setText(text);
        }
      } else if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
          const table = element.asTable();
          const rows = table.getNumRows();
          const cols = table.getNumColumns();

          for (let row = 0; row < rows; row++) {
            for (let col = 0; col < cols; col++) {
              const cell = table.getCell(row, col);
              if (cell.getText()) {
                let cellText = cell.getText().asString();
                for (const variableName in variableValues) {
                  const variablePattern = `\\$\\{${variableName}\\}`;
                  const regex = new RegExp(variablePattern, 'g');
                  cellText = cellText.replace(regex, variableValues[variableName]);
                }
                cell.getText().setText(cellText);
              }
            }
          }
      }
    });
  });

  // Export the presentation as a PDF
  const pdfBlob = presentation.getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdfBlob).setName(presentationName + '.pdf');

  // Optional: open the PDF in a new tab
  DocsList.getFileById(pdfFile.getId()).getViewerUrl();

  // Optional: delete the temporary file after opening.
  // DriveApp.getFileById(pdfFile.getId()).setTrashed(true);
}

/**
 * Extracts variable names from a list of slides.
 *
 * @param {Slide[]} slides The slides to extract variable names from.
 * @return {string[]} An array of unique variable names.
 */
function extractVariableNames(slides) {
  const variableNames = new Set();

  slides.forEach(slide => {
    const pageElements = slide.getPageElements();

    pageElements.forEach(element => {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        if (shape.getText()) {
          const text = shape.getText().asString();
          const regex = /\$\{(.*?)\}/g; // Matches ${variable_name}
          let match;

          while ((match = regex.exec(text)) !== null) {
            variableNames.add(match[1]); // Extract the variable name
          }
        }
      } else if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
          const table = element.asTable();
          const rows = table.getNumRows();
          const cols = table.getNumColumns();

          for (let row = 0; row < rows; row++) {
            for (let col = 0; col < cols; col++) {
              const cell = table.getCell(row, col);
              if (cell.getText()) {
                const text = cell.getText().asString();
                const regex = /\$\{(.*?)\}/g;
                let match;
                while ((match = regex.exec(text)) !== null) {
                  variableNames.add(match[1]);
                }
              }
            }
          }
      }
    });
  });

  return Array.from(variableNames);
}


function onOpen() {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem('Replace Variables and Export', 'replaceVariablesAndExportPdfWithForm') // Add a menu item
    .addToUi();
}