// Serves the Web App UI
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestão Financeira')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Utility to generate unique short IDs
function generateId(prefix) {
  return prefix + '-' + Utilities.getUuid().substring(0, 8);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}