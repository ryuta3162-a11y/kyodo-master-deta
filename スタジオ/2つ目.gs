/**
 * ウェブアプリとしてアクセスされた時に実行される関数
 * index.html を読み込んで画面に表示します。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('スタジオ集計ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}