# 登壇デモ 一式 — pull したらまずこれ

静岡ママヨロサロン登壇用。中身は3つ：**(A) 通しトーク / (B) pptx生成 3段階デモ / (C) 無料相談フォーム**。

---

## A. 通しトーク
`shizuoka_talk_v2.html` をブラウザで開く。`→ / Space`（またはクリッカー）で1画面ずつ進む。
流れ：イントロ → 川（点→引き→線）→ 種明かし → ①②③ → CTA。
川ステージは `→` を数回ぶん押すと内部の演出（詰まり→点→引き→線）が進み、そのあと次のセクションへ。
※ v1（`shizuoka_talk.html`）は旧版。本番は **v2**。

## B. pptx生成 実践デモ（3段階・同じ3名）
題材：8月生まれ **佐藤 健一(8/5)・鈴木 美咲(8/14)・田中 大輔(8/27)**。

- **① 手動** … `demo_3stage/01_手動_誕生日スライド.pptx` を開いて見せる（担当者が手作りした体。素朴・時間がかかる）
- **② 点（AIでその場生成）** … `demo_3stage/02_点_AI生成.pptx` を開いて見せる（AIに頼んだ1枚。速くて綺麗。ただし毎回ブレる・要プロンプト/目視確認・DBと無関係な単発）
- **③ 線（プログラム×DB）** … 下記を実行 → 同じ3名が **DBから自動で完成**（プロンプト不要・毎回同品質・一瞬・ほぼ無料）

### ③ の実行（ノートPC）
Python が必要。
```
cd demo_slidegen
pip install -r requirements.txt    # 初回のみ（python-pptx, Pillow）
python generate_slide.py           # 8月のお祝いを生成 → PowerPointが自動で開く
```
- **連続生成（月を変えて魅せる）**：`python generate_slide.py -i` → `8`↵ / `11`↵ / `3`↵ …
  - 11月＝誕生日1・記念日1、10月＝非公開1名スキップ、など「DBから拾って一瞬」を実演できる
- データ：`demo_slidegen/db/members.json`（会員8名）・`event.json`（例会＝2026-08）。**8月＝①②と同じ3名**。
- フォント：**游明朝／游ゴシック**（Windows標準）前提。
- オプション：`--month 11`（月指定）／`--no-open`（開かない）／`--rebuild-bg`（背景作り直し）

## C. 無料相談フォーム
`salon_form/`（`index.html` / `Code.gs` / `README.md`）。talk の CTA・QR の差し込み先。
**デプロイ＆スプレッド接続は未実施。** 当日前に `salon_form/README.md` の手順（約10分）で
スプレッド＋Apps Script を作成 → 出てきたURLを `index.html` の `ENDPOINT` に貼る → 公開 → QR化。
（フォームは「無料相談＝ホット」「案内・資料＝ソフト」の二段構え。回答は同じシートに `種別` で集約）

---

## メモ / 注意
- CC のコミットで「`.git/index.lock` が既にある」と出たら、その 1 ファイルを削除してから再実行。
- 生成物（`demo_slidegen/out/`・`demo_slidegen/assets/bg.png`）は実行時に自動生成されるので、リポジトリに無くてOK。
- 同意リンク（プライバシーポリシー）のURLは本番前に差し込み（`salon_form/index.html` 内のリンク）。
