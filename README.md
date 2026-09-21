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
全OS版ともポータブルです。macOS版は署名・公証済みDMGを開いて`.app`をApplicationsへドラッグし、アンインストール時はその`.app`をゴミ箱へ移してください。Windows／Ubuntu版は展開したフォルダを削除すればアンインストールできます。<br>
**実行ファイルを使用する場合**
<br>
- Windows : **TtT_ScoreInput.exe**<br>
- Ubuntu : **TtT_ScoreInput**<br>
- macOS : **TtT_ScoreInput**<br>

**TtT_ScoreInput.pyから実行する場合**
<br>
Python 3.14をインストールし、以下のコマンドで依存関係を導入します。<br>
```
python -m pip install -r requirements.txt
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
| Windows 11 x64 | 対応（Windows Server 2025でビルド） |
| Ubuntu latest x64 | 対応 |
| macOS 26 Apple Silicon | 対応 |

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
Python 3.14をインストールし、任意のディレクトリで以下を実行します。macOSのHomebrew版Pythonを使う場合は、先に`brew install python-tk@3.14`でTkも導入してください。

```
git clone https://github.com/ryuya0124/TtT_ScoreInput.git
cd TtT_ScoreInput
python -m pip install -r requirements-build.txt
python -m unittest discover -s tests -v
python -m PyInstaller --clean --noconfirm --windowed --onedir --name TtT_ScoreInput TtT_ScoreInput.py
```

GitHub Actionsはタグをpushすると3 OS向け成果物を作成し、GitHub Releaseへ自動添付します。macOS版はfastlane matchで管理するDeveloper ID Application証明書で署名し、Appleの公証とstapleを完了したDMGとして配布します。Windows版は誤検知リスクを下げるため、公式PyInstaller 6.22.3のbootloaderをWindowsランナー上でソースからビルドし、`onedir`形式で配布します。これは誤検知の完全な防止を保証するものではありません。

# 注意事項
このツールは非公式のものであり、使用にあたっては自己責任でお願いします。<br>
ツールの使用によって生じた損害や問題について、作成者は一切の責任を負いません。<br>
使用前に必ず内容を理解し、納得の上でご利用ください。<br><br>
二次配布やフォークして改良などはご自由にどうぞ！

# 作成者
りゅうや<br>
