# 経堂利用人数（Git フォルダ名）

## Google ブック名（正式）

**経堂利用人数集計　データ　FW　ランニングマシン**

| 項目 | 値 |
|------|-----|
| URL | https://docs.google.com/spreadsheets/d/1qWOf5apYTULT3YWw8oF2ZQ8p73jBM4-7Y7NY8adcsXg/edit?gid=363661640#gid=363661640 |
| ID | `1qWOf5apYTULT3YWw8oF2ZQ8p73jBM4-7Y7NY8adcsXg` |
| メニュー | **ジム施設集計システム** |

## GAS をどこに貼るか

```
経堂利用人数/
├── 元データ/
│   ├── Code.gs          ← ★ まずここにブックから貼る（正本・凍結）
│   └── README.md        ← 手順詳細
├── Code.gs              ← 作業用（元データと同内容を維持）
└── （連携はアイデアプール側 SyncFromUsageCount.gs）
```

| やりたいこと | 貼る場所 |
|--------------|----------|
| ブックの GAS を Git に保存 | **[元データ/Code.gs](元データ/Code.gs)** |
| 作業・改修のベース | [Code.gs](Code.gs) |
| アイデアプールへリアルタイム表示 | [一元管理/SyncFromUsageCount.gs](../一元管理/SyncFromUsageCount.gs) |

**Google ブック側** … 拡張機能 → Apps Script → `Code.gs`（Git と同期するとき）

## 中身

- MAX人数入力（PWAシミュレーション予測値セット済み）
- 時間帯別（0〜4時 … 21〜0時）・日次・月平均
- シート: `FW_データ入力` / `TM_データ入力` / 各分析ダッシュボード

## 関連

- 台帳: [マスタ/スプレッドシート一覧.md](../マスタ/スプレッドシート一覧.md)
- 保管ルール: [マスタ/元データ保管ルール.md](../マスタ/元データ保管ルール.md)
