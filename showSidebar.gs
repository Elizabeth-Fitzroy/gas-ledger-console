/**
 * 操作用サイドバーを表示します。
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('セット組＆棚卸 Console')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}
