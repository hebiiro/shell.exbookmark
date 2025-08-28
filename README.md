# 🎉EXブックマーク

* エクスプローラの右クリックメニューにブックマークを追加できるようにします。

> [!TIP]
> このブックマークはaviutl/aviutl2のエクスプローラやファイル選択ダイアログだけでなく、その他のアプリとも共有できる可能性があります。

## 🚀インストールまたは🔥アンインストール

1. `exbookmark_manager.exe`を実行してください。
1. `exbookmarkの管理`ダイアログが表示されます。
1. `インストール`または`アンインストール`ボタンを押してください。

> [!IMPORTANT]
> 管理者権限でレジストリを変更するので、OSなどから警告が出る場合があります。

### 🔧レジストリ

* `exbookmark_manager.exe`でアンインストールできない場合は、以下のレジストリを削除してください。

```
	HKEY_CLASSES_ROOT\*\shellex\ContextMenuHandlers\exbookmark
	HKEY_CLASSES_ROOT\Folder\shellex\ContextMenuHandlers\exbookmark
	HKEY_CLASSES_ROOT\Directory\Background\shellex\ContextMenuHandlers\exbookmark
```

> [!NOTE]
> すべての場所の右クリックメニューを拡張しようとするとレジストリへの書き込み箇所が増えてしまうため、上記実用最小限の範囲だけに厳選しています。

## 💡使い方

### 🏷️ブックマークを使用する

1. エクスプローラなどで右クリックします。
1. 右クリックメニューでブックマーク項目を選択します。

### 🏷️ブックマークを追加(または削除)する ※簡易

1. エクスプローラなどで右クリックします。
1. 右クリックメニューで`この階層に現在のパスを追加(または削除)`を選択します。
	* ファイル(またはフォルダ)を選択して右クリックした場合は、選択ファイル(またはフォルダ)が対象になります。
	* バックグラウンドで右クリックした場合は、現在開いているフォルダが対象になります。

### 🏷️ブックマークを編集する ※高度

1. `exbookmark/folder_tree.xml`をテキストエディタで編集します。

> [!IMPORTANT]
> xmlの文法エラーがあると読み込みに失敗するので、注意して編集してください。

## 📝folder_tree.xml

* `node`✏️メニュー項目です。入れ子にできます。
	* `label`✏️メニュー項目の表示名です。
	* `path`✏️メニュー項目をクリックしたときに選択または実行されるパスです。
		* パスは親項目からの相対パスと解釈され結合されます。
	* `icon_path`✏️アイコンのパスです。
		* 表示アイコンを変更したい場合に指定します。
		* 指定しなかった場合は`path`が使用されます。
	* `icon_index`✏️アイコンのインデックスです。
		* 表示アイコンを変更したい場合に指定します。
		* 指定しなかった場合は`icon_path`自体のアイコンが使用されます。
	* `absolute`✏️`path`を絶対パスとして解釈するフラグです。
		* 指定した場合は`path`は親項目のパスと結合されません。
	* `execute`✏️パスを選択するのではなく、実行するフラグです。
		* 指定した場合はパスが直接実行されます。
	* `type`✏️項目のタイプです。
		* `separator`✏️項目をセパレータにします。
		* `add`✏️項目をパスを追加するコマンドにします。
		* `remove`✏️項目をパスを削除するコマンドにします。

### 🏷️例：ユーザーフォルダ

```xml
	<!-- 子ノードはabsolute=""を指定していないので、親ノードからの相対パスになります。 -->
	<node label="ユーザー" path="%USERPROFILE%" absolute="">
		<node label="ダウンロード" path="Downloads"/>
		<node label="ドキュメント" path="Documents"/>
		<node label="ビデオ" path="Videos"/>
		<node label="ミュージック" path="Music"/>
		<node label="ピクチャ" path="Pictures"/>
		<node label="リポジトリ" path="source\repos"/>
	</node>
```

### 🏷️例：ファイルを選択する

```xml
	<node label="動画出力のテスト" path="C:\my\test.mp4" absolute=""/>
```

### 🏷️例：ファイルを開く

```xml
	<node label="メモ" path="C:\my\memo.txt" absolute="" execute=""/>
```

### 🏷️例：電卓を実行する

```xml
	<node label="電卓" path="calc.exe" absolute="" execute=""/>
```

### 🏷️例：ウェブサイトを開く

```xml
	<node label="AviUtlのお部屋" path="https://spring-fragrance.mints.ne.jp/aviutl/" icon_path="shell32.dll" icon_index="135" absolute="" execute=""/>
```

### 🏷️例：簡易的に追加/削除できるようにする

```xml
	<node type="separator"/>
	<node type="add" label="この階層に現在のパスを追加" icon_path="shell32.dll" icon_index="294"/>
	<node type="remove" label="この階層から現在のパスを削除" icon_path="shell32.dll" icon_index="131"/>
```

### 🏷️エクスプローラのテスト

* 開発者用です。インプロセスの簡易エクスプローラを表示します。

### 🏷️ファイル選択ダイアログのテスト

* 開発者用です。ファイル選択ダイアログを表示します。

## 🔖更新履歴

* 🔖r1 #2025年08月22日
	* 🎉初版

## ⚗️動作確認

* Win11 Home 24H2

## 👽️作成者情報
 
* 作成者 - 蛇色 (へびいろ)
* GitHub - https://github.com/hebiiro
* X - https://x.com/io_hebiiro

## 🚨免責事項

> [!IMPORTANT]
> この作成物および同梱物を使用したことによって生じたすべての障害・損害・不具合等に関しては、私と私の関係者および私の所属するいかなる団体・組織とも、一切の責任を負いません。各自の責任においてご使用ください。
