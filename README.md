#  請求書・見積書 自動作成ツール

## ・ 概要

本ツールは、顧客からの注文情報と商品マスタをもとに、Excel形式の見積書・請求書を自動生成します。  

Pythonを使用することで、膨大で煩雑なPC作業の自動化・高速化が可能になります。また、正確性が高いうえに、何度でも再現が可能です。

---

## ・ 使用環境

- OS：Windows 10 / 11
- Python環境：vertualenv環境
- 実行形式：`.py`（Pythonスクリプト） 

---

## ・ フォルダ構成

<pre lang="markdown">05_invoice_generator/
├── input/ # 注文情報・商品マスタ（CSV）
│ ├── orders.csv
│ └── products.csv
│
├── output/ # 出力帳票（Excelファイル）
│ └── O0001_quotation.xlsx など
│
├── template/ # 帳票テンプレート（Excel）
│ └── quotation_template.xlsx
│
├── script/ # スクリプト本体
│ ├── quotation_generator.py
│ ├── invoice_generator.py
│ └── common_utils.py
│
└── README.md</pre>


> ※ script/に、開発段階で使用した.ipynbファイルを開発ログとして保存してあります。

---

## ・ 出力例

![見積書サンプル](sample_images/O0003_quotation.png)
![請求書サンプル](sample_images/O0003_invoice.png)

> ※ 実際の出力は `template/quotation_template.xlsx` をベースにしています

---

## ・ ライセンス

本プロジェクトは [MITライセンス](LICENSE) のもとで公開されています。  
商用・非商用を問わずご自由にご利用いただけますが、再配布時はクレジットの記載をお願いします。
