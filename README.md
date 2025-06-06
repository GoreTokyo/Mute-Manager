# Mute Manager

Mute Managerは、Windowsのマイクのミュート状態を管理するための簡単なコンソールアプリケーションです。このプログラムは、ユーザーがマイクのミュート/ミュート解除を操作したり、現在のミュート状態を確認したりすることができます。

## 特徴

- マイクのミュートとミュート解除を簡単に行える
- 現在のマイクのミュート状態を確認できる
- 操作履歴をログファイル（`MuteManager.log`）に記録
- 日時とともに操作メッセージを表示

## 必要条件

このプログラムを実行するには、次の要件があります：

- Windowsオペレーティングシステム（Windows 7以降を推奨）
- Visual Studio などのC++開発環境
- COMライブラリのサポート（Windows標準）

## 使用方法

1. ソースコードをVisual StudioなどのIDEでビルドします。

2. プログラムを実行すると、以下のメニューが表示されます：

`1: ミュート 2: ミュート解除 3: ミュート状態の確認 4: 終了 操作を選択してください:`

3. メニューから番号を入力して、次の操作を選択します：
- `1`: マイクをミュート
- `2`: マイクのミュートを解除
- `3`: 現在のミュート状態を確認
- `4`: プログラムを終了

4. 操作の結果は、画面に表示され、同時にログファイルに記録されます。ログファイルは`MuteManager.log`という名前でプログラムが実行されているディレクトリに作成されます。

## ログファイルの内容

操作の結果は、`MuteManager.log`に記録されます。

## 詳細

### クラスと関数

- **ComInitializer**:
  - COMライブラリを初期化および終了するためのクラス。
  
- **GetAudioEndpointVolume**:
  - オーディオデバイスのボリュームインターフェースを取得します。

- **CheckMuteStatus**:
  - マイクのミュート状態を確認し、結果を表示します。

- **MuteMicrophone**:
  - マイクをミュートにします。

- **UnmuteMicrophone**:
  - マイクのミュートを解除します。

- **LogMessage**:
  - ログファイルにメッセージを記録します。

- **PrintAndLogMessage**:
  - 画面にメッセージを表示し、ログファイルに記録します。

## 注意事項

- 本プログラムは、マイクのミュート状態を管理するための簡単なツールです。
- プログラムが正常に動作するには、管理者権限での実行が必要な場合があります。
- ログファイルは、プログラムが実行されているディレクトリに保存されます。

## ライセンス

このプログラムは、MITライセンスの下で提供されます。自由に使用、変更、配布することができます。
