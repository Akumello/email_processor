/**
 * Serves the web app HTML.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The rendered index page
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes an HTML partial file (CSS/JS) in a template.
 * Usage in HTML: <?!= include('stylesheet') ?>
 * @param {string} filename - The name of the HTML file to include (without .html)
 * @returns {string} The file contents as a string
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
