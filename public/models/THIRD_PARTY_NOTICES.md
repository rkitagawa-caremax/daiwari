# 端末内AIに含まれる第三者ソフトウェア・モデル

台割くんの「AIアシスト」は、外部AI APIへ商品データを送らず、以下のモデルと実行環境を利用者のブラウザ内で実行します。

- Ruri v3 30m / Ruri v3 30m ONNX — Apache License 2.0
  - https://huggingface.co/cl-nagoya/ruri-v3-30m
  - https://huggingface.co/onnx-community/ruri-v3-30m-ONNX
- Transformers.js — Apache License 2.0
  - https://github.com/huggingface/transformers.js
- ONNX Runtime — MIT License
  - https://github.com/microsoft/onnxruntime

モデルは日本語の商品テキストを数値ベクトルへ変換する用途に限って使います。モデル、トークナイザー、WASM実行ファイルはこのアプリと同じ配信元から読み込みます。
