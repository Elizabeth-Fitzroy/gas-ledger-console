/**
 * @fileoverview スプレッドシート起動時のカスタムメニュー統合設定
 */

/**
 * スプレッドシートを開いた時にカスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. 書籍検索・API自動補完メニュー
  ui.createMenu('★書籍検索メニュー')
    .addItem('1. OpenBD から取得 (文庫・新刊向き)', 'updateUsingOpenBD')
    .addItem('2. Google Books から取得 (網羅性重視)', 'updateUsingGoogleBooks')
    .addItem('3. NDL から取得 (図書館データ・古い本向き)', 'updateUsingNDL')
    .addSeparator()
    .addItem('全APIを順番に試す (自動フォールバック)', 'updateLookupBookIntegrated')
    .addToUi();

  // 2. 販売・出品コンソール（旧機能）
  ui.createMenu('📚 在庫管理')
    .addItem('販売・出品コンソールを開く', 'showSalesConsoleSidebar')
    .addToUi();

  // 3. セット組・棚卸コンソール（新機能）
  ui.createMenu('★セット組・棚卸')
    .addItem('右側サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('データ抽出（フェッチ）', 'fetchToConsoleFromSheet')
    .addItem('ライトバック（反映）', 'writeBackToLedger')
    .addToUi();
}