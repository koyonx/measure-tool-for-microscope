# Measure Tool for Microscope

顕微鏡画像の距離・面積を測定するためのPythonツールです。

## 必要な環境

- Python 3.x
- OpenCV (`cv2`)
- NumPy
- openpyxl（`measure_and_excel.py`のみ）

### インストール

```bash
pip install opencv-python numpy openpyxl
```

## ツール一覧

### 1. measure_tool.py

基本的な測定ツール。画像上で距離と面積を測定できます。

#### 使い方

```bash
python measure_tool.py <画像ファイル名> [スケール値(μm)]
```

#### 例

```bash
# スケールバーが50μmの場合（デフォルト）
python measure_tool.py sample.jpg

# スケールバーが100μmの場合
python measure_tool.py sample.jpg 100
```

---

### 2. measure_and_excel.py

測定結果をExcelファイルに自動保存する拡張版です。測定時にメモを入力でき、データを整理して保存します。

#### 使い方

```bash
python measure_and_excel.py <画像パス> [スケール値] [Excelファイル名]
```

#### 例

```bash
# デフォルト設定（スケール50μm、結果はmeasure_results.xlsxに保存）
python measure_and_excel.py sample.jpg

# スケール100μm、カスタムExcelファイル名
python measure_and_excel.py sample.jpg 100 my_results.xlsx
```

#### Excelの出力形式

3つのシートが自動作成されます：

| シート名 | 内容 |
|----------|------|
| All_Data | 全ての測定データ |
| Distance_Only | 距離測定データのみ |
| Area_Only | 面積測定データのみ |

---

## 操作方法（共通）

### ステップ1: スケールキャリブレーション

1. ツールを起動すると画像が表示されます
2. 画像内のスケールバーの**左端**と**右端**をクリック
3. キャリブレーションが完了し、測定モードに移行します

### ステップ2: 測定

| キー | 機能 |
|------|------|
| `d` | 距離測定（2点をクリック後に押す） |
| `a` | 面積測定（3点以上で囲んだ後に押す） |
| `r` | 画面リセット |
| `q` | 終了 |

### 測定手順

**距離測定:**
1. 測定したい2点をクリック
2. `d`キーを押す
3. 結果が画面に表示される（Excel版はメモ入力後に保存）

**面積測定:**
1. 測定したい領域を3点以上でクリックして囲む
2. `a`キーを押す
3. 結果が画面に表示される（Excel版はメモ入力後に保存）

---

## 注意事項

- スケールバーの長さを正確にクリックしないと、測定結果に誤差が生じます
- `measure_and_excel.py`使用時、Excelファイルを開いたまま保存しようとするとエラーになります
- 画像ファイルが見つからない場合はエラーメッセージが表示されます
