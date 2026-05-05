# VoxCord
<img width="2344" height="1563" alt="logo" src="https://github.com/user-attachments/assets/f74c6e08-798d-488a-9a58-7040f66acf8b" />

Discord のメッセージを VoiceVox で読み上げる、PySide6 製のGUI付き TTS ボット管理アプリです。

## 主な機能

* Discord Bot 接続
* 指定した Text Channel → Voice Channel の読み上げ
* メッセージ整形

  * URL を「URL」に置換
  * 画像添付を「画像」として読み上げ
  * 置換ルール対応
  * メンションやカスタム表現の整形
* 話者ごとの個別割り当て
* 読み上げ速度の変更
* TTS の開始・停止
* Webhook 通知

  * TTS を開始しました
  * TTS を終了します
  * エラーが発生しました
* GUI で設定編集
* ログ表示
* AppData へのログ保存

## 動作環境

* Windows 10 / 11
* Python 3.13 系
* Discord Bot が利用できる環境
* VoiceVox と FFmpeg を利用できる環境

## VOICEVOX使用上の注意

* このアプリは著作権の関係からVOICEVOXをレポジトリに含めていません。
* インストール時に自動でダウンロードされ、配置されます。

## フォルダ構成(完成系ですので自分で動かしてくださいね)

```text
VoxCord/
├─ main.py
├─ gui.py
├─ discord_service.py
├─ config_manager.py
├─ tts_engine.py
├─ message_processor.py
├─ config.json
├─ speakers.json
├─ logo.ico
├─ logo.png
├─ assets/
├─ FFmpeg/
|　　└─ffmpeg.exeなど
└─ VOICEVOX/
　　└─run.exeなど
```

## 外部に置くもの

アプリ本体とは別に、実行ファイルと同じ場所に以下を置きます。(インストール時は自動で配置)

* `assets/`
* `FFmpeg/`
* `VOICEVOX/`
* `speakers.json`

設定ファイルとログは AppData に保存されます。

```text
C:\Users\<ユーザー名>\AppData\Local\VoxCord\
├─ config.json
├─ logs\
└─ temp\
```

## 使い方

1. GUIで`config.json` に Discord Bot トークンを設定します。
2. GUIで`channel_pairs` に Text Channel ID と Voice Channel ID を設定します。
3. 必要に応じて `member_voice_map` と `replace_rules` を編集します。
4. `VoxCord.exe` または `main.py` を起動します。
5. GUI から Start を押して TTS を開始します。

## 設定項目

### bot_token

Discord Bot のトークン。

### default_speaker

未登録ユーザーに使う既定話者 ID。

### speed

読み上げ速度。

### channel_pairs

読み上げ対象のチャンネル対応表。

```json
{
  "guild_id": "123456789012345678",
  "text_channel_id": "123456789012345678",
  "voice_channel_id": "123456789012345678",
  "enabled": true
}
```

### member_voice_map

ユーザーごとの話者割り当て。

```json
{
  "123456789012345678": {
    "speaker_id": 2,
    "enabled": true
  }
}
```

### replace_rules

読み上げ前の置換ルール。

```json
{
  "w": "わら",
  "www": "わらわら"
}
```

### filters

メッセージフィルタ。

* `ignore_bots`
* `ignore_commands`
* `command_prefix`
* `max_length`
* `ignore_empty`

### queue

再生キュー設定。

* `max_size`
* `drop_old_when_full`

## GUI

GUI では以下を操作できます。

* Bot Token の保存
* チャンネルペアの追加・削除・有効無効切り替え
* 話者マッピングの編集
* 置換ルールの編集
* 読み上げ速度の変更
* 設定のインポート / エクスポート
* ログ表示

## 話者一覧

話者情報は `speakers.json` から読み込みます。
1 行 JSON で保存されている形式でも読み取れるようにしています。

例:

```json
[{"name":"四国めたん","speaker_uuid":"...","styles":[{"name":"ノーマル","id":2,"type":"talk"}]}]
```

## ログ

ログは以下に保存されます。

```text
C:\Users\<ユーザー名>\AppData\Local\VoxCord\logs\latest.log
```

## 初回起動時の注意

* `config.json` がなければ自動生成されます。
* `FFmpeg`と`VoiceVox`はインストール時にダウンロード、展開されます。
* Discord の音声機能には `PyNaCl` と `opus.dll` が必要です。

## インストーラー化

Program Files に展開する場合は Inno Setup などでインストーラーを作成できます。
設定とログは AppData に保存されるため、Program Files 配下でも安全に動作します。

## トラブルシュート

### 音声が再生されない

* `PyNaCl` がインストールされているか確認
* `opus.dll` が 64bit 版か確認
* `FFmpeg` が存在するか確認

### VOICEVOX が起動しない

* `VOICEVOX\run.exe` があるか確認
* パスが正しいか確認

### ログが出ない

* `AppData\Local\VoxCord\logs\latest.log` を確認
* exe 直下ではなく AppData 側を確認

## ライセンス

このプロジェクトが利用する外部ソフトウェアのライセンスに従ってください。
Discord.py、PySide6、VoiceVox、FFmpeg などの各ライセンスも確認してください。

## 開発メモ

* 設定は AppData 側の `config.json` を使用
* ログも AppData 側へ保存
* 外部フォルダは exe と同じ場所から参照
* 音声再生では `opus.dll` が必要

This software includes components from 7-Zip.

7-Zip is licensed under the GNU Lesser General Public License (LGPL).

Copyright (C) 1999-2024 Igor Pavlov

You can obtain the source code of 7-Zip from:
https://www.7-zip.org/

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY.

This software uses libopus.

Copyright (c) 2011-2024 Xiph.Org Foundation

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

- Redistributions must retain the above copyright notice,
  this list of conditions and the following disclaimer.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES ARE DISCLAIMED.

