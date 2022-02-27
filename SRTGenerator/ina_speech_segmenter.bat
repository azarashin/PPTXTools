rem 仮想環境のactivate(必要に応じて書き換えてください)
call conda activate voice_text

python ina_speech_segmenter.py %1 %2

rem 仮想環境のdeactivate(必要に応じて書き換えてください)
conda deactivate
