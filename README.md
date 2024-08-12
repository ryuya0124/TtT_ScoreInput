# TtT_ScoreInput
# 説明
テトテコネクトの全楽曲と難易度をまとめたExcelファイルに、TetoteConnect-Scoreを使用してダウンロードしたCSVファイルを元にプレイ状況を書き込むツールです。<br>
<br><br>
Excelファイルはこちらから利用できます。<br>
[TetoteConnect-ClearSheet](https://github.com/neco0814/TetoteConnect-ClearSheet)<br>
CSVファイルはこちらから利用できます。<br>
[TetoteConnect-Score](https://github.com/3-show/TetoteConnect-Score)<br>
[TetoconeScoreDataTool](https://github.com/chespins/TetoconeScoreDataTool)

![image](https://github.com/user-attachments/assets/b2385a2c-a2f1-4b67-828b-dacd442af340)
![image](https://github.com/user-attachments/assets/75befec7-3315-43dc-b1b1-adf3b520aaf9)
![image](https://github.com/user-attachments/assets/423f1f4d-5305-4c28-9039-4b5d06d3ffb2)

# Excel表記
| 表記 | 意味 |
----|----
| AP | All Perfect |
| FC | Full Combo |
| CL | Clear |
| FL | Failed |

# 使用方法
[Release](https://github.com/ryuya0124/TtT_ScoreInput/releases)から環境に合わせたファイルをダウンロードします。<br>
**実行ファイルを使用する場合**
<br>
- Windows : **TtT_ScoreInput.exe**<br>
- Ubuntu : **TtT_ScoreInput**<br>
- macOS : **TtT_ScoreInput**<br>

**TtT_ScoreInput.pyから実行する場合**
<br>
以下のコマンドでモジュールをインストールする必要があります。<br>
```
pip install pandas openpyxl
```
<br>
<br>
それぞれのファイルパスを設定します。<br>

**デフォルトに設定** を押すとスクリプトのフォルダ内にあるExcelファイルとCSVファイルを自動設定します。<br>
> [!TIP] 
> Excelファイルの名前は**TtT_ClearSheet.xlsx** である必要があります。

 **処理を開始** を押すと処理が開始されます。<br>

> [!WARNING] 
> **TtT_ClearSheet.xlsx** に上書き保存されます。

# 対応OS
| OS | 対応状況 |
----|----
| Windows11 | 対応 |
| Windows10 | 多分動く！ |
| Ubuntu | デバッカー求む! |
| macOS 13,14 | 修正作業中 |

<br>

> [!WARNING]
> WSLで実行する場合は以下の方法でフォントの修正が必要です。<br>
> https://nexem.hatenablog.com/entry/2020/07/18/223540
<br>

# 対応言語
| 言語 | 対応状況 | 
----|----
| 日本語 | 対応 |
| 英語 | 気が向いたら |
| その他 | 気が向いたら |

# ビルド方法
任意のディレクトリで以下を実行します。

- ビルド準備
```
git clone `https://github.com/ryuya0124/TtT_ScoreInput.git`
cd TtT_ScoreInput
```

- 実行ファイルに変換するモジュールのインストール
通常のPyinstallerではWindows Defenderにウイルス検知されてしまうので、以下のものを使用します。
```
pip install git+https://github.com/fa0311/pyinstaller
```

- ビルド開始
```
pyinstaller --noconsole TtT_ScoreInput.py
```

**buildフォルダ**は削除してください。
**distフォルダ**の中に実行ファイルがあります。

# 注意事項
このツールは非公式のものであり、使用にあたっては自己責任でお願いします。<br>
ツールの使用によって生じた損害や問題について、作成者は一切の責任を負いません。<br>
使用前に必ず内容を理解し、納得の上でご利用ください。

# 作成者
りゅうや<br>
Twitter(𝕏) : [@_ryuya_0124](https://x.com/_ryuya_0124)
