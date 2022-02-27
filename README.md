# PPTXTools

## 前提条件

- PowerPoint をインストールしておいてください
  - PowerPoint 2016 にて動作確認しています
- python をインストールしておいてください
  - 必要に応じて仮想環境を構築しておいてください
  - python3.10 にて動作確認済みです
- ffmpeg をインストールしておいてください
  - パスを通してffmpeg コマンドを実行できる状態にする必要があります

## 環境構築手順

### python モジュールの追加

下記を実行し、python のinaSpeechSegmenter モジュールを追加しておいてください

```
pip install inaSpeechSegmenter
```

### .bat ファイルの書き換え

使用しているpython の仮想環境に応じて、

```
SRTGenerator/python/ina_speech_segmenter.bat
```

を書き換えてください。

本ファイルの中身は下記のようになっています。



```
rem 仮想環境のactivate(必要に応じて書き換えてください)
call conda activate voice_text

python ina_speech_segmenter.py %1 %2

rem 仮想環境のdeactivate(必要に応じて書き換えてください)
conda deactivate

```

同梱されている状態ではconda による仮想環境を使用しています（仮想環境名はvoice_text）。

お使いの環境に応じて仮想環境のactivate/deactivate 部分を書き換えてください。




## ツールの概要＆使い方

doc/Promotion.pptx を参照してください

