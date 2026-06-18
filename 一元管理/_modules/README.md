# 開発用モジュール（Apps Script には貼らない）

編集はここで行い、**`../Code.gs` に統合してから** アイデアプールへ貼ります。

| ファイル | 内容 |
|----------|------|
| `SyncFromReformer.gs` | リフォーマー取り込み |
| `ReformerDashboardAggregate.gs` | deta 集計 ⑦⑧⑨ |
| `SyncFromUsageCount.gs` | 利用状況 FW/TM |

## Code.gs への統合

`_modules` を更新したあと、Cursor または手動で `Code.gs` を再生成してください。  
通常は **`Code.gs` を直接編集** し、必要に応じてここへコピーバックしても構いません。

**Apps Script に置くのは `Code.gs` 1ファイルだけ** です。
