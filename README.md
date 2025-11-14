# python-pptx Demo

Python-pptx を使用した PowerPoint プレゼンテーション自動生成のデモプロジェクト

## 機能

- **テーブル** - データを見やすい表形式で表示
- **棒グラフ** - カテゴリ別の比較に最適
- **折れ線グラフ** - 時系列データの推移を可視化
- **円グラフ** - 割合や構成比を表示
- **帯グラフ** - 100%積み上げ横棒グラフ（横向き/縦向き）
- **箇条書き** - 階層構造のテキスト表示
- **2 カラムレイアウト** - メリット/デメリットなどの比較表示
- **カスタマイズ可能なスタイル** - 統一されたブランドカラーとフォント

### スクリーンショット

[![Image from Gyazo](https://i.gyazo.com/6b9c3357e255d8901a7f2b7781bfcaf9.png)](https://gyazo.com/6b9c3357e255d8901a7f2b7781bfcaf9)

## 必要要件

- Python 3.7+
- python-pptx
- Pillow

## インストール

```bash
# リポジトリのクローン
git clone git@github.com:pikum99/python-pptx.git
cd python-pptx

# 仮想環境の作成
python3 -m venv venv
source venv/bin/activate  # Windowsの場合: venv\Scripts\activate

# 依存パッケージのインストール
pip install python-pptx Pillow
```

## 使い方

```bash
python demo.py
```

実行すると、`demo_presentation.pptx` が生成されます。

## カスタマイズ

### スタイル定数の変更

`demo.py` の冒頭で定義されている定数を変更することで、スタイルをカスタマイズできます：

```python
TITLE_FONT_SIZE = Pt(32)          # タイトルのフォントサイズ
CHART_Y_POSITION = Inches(1.2)    # グラフのY位置
CHART_HEIGHT = Inches(4.2)        # グラフの高さ
BRAND_COLOR = RGBColor(0, 51, 102)  # ブランドカラー
```

### データの変更

`main()` 関数内の `product_data` と `time_series_data` を編集することで、表示するデータを変更できます。

## 生成されるスライド

1. タイトルスライド
2. アジェンダ（目次）
3. データテーブル
4. 売上比較（棒グラフ）
5. 月次推移（折れ線グラフ）
6. 市場シェア（円グラフ）
7. 評価比較（帯グラフ - 横向き）
8. 評価比較（縦向き積み上げグラフ）
9. 製品評価（2 カラムレイアウト）
10. アクションプラン（階層的箇条書き）
11. まとめスライド

## 参考リンク

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
