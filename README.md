# ExToolsForExcelTest
Excelでテストを管理する際の、便利ツールです。<br>
[私のブログ](https://soncho-works.com/2022/06/04/excel-test-tool/)にもう少し詳しく書いてあるので、そっちを見てもらってもいいと思います。

# できること
- 開発したアプリにフォーカスしたまま、Excelにテスト結果を残せます。<br>アプリとExcelをマウスで行き来したり、Excelの所定のセルを探したりする必要はありません。
- OKやらNGやらをタイプする必要もありません。
- 次にテストするべき項目が一番上の行に表示されるので、テスト項目の見間違えを防げます。
エビデンス(キャプチャ画像)も自動で取ってExcelに保存してくれます。

## 機能
- Alt+9で、Excelの指定の列にOKと書いて、次の行にフォーカス移動&次の行を一番上にスクロールします。
- Ctrl+Alt+9で、Excelの指定の列にNGと書いて、次の行にフォーカス移動&次の行を一番上にスクロールします。
- Alt+0で、1行スキップしてくれます。
- Alt+8で、Alt+9の動作に加えて、アクティブなウィンドウのスクリーンショットも指定のシートに残します。
- Ctrl+Alt+8で、Ctrl+Alt+9の動作に加えて、アクティブなウィンドウのスクリーンショットも指定のシートに残します。
- これらのショートカットは、フォーカスの当たっているアプリに関係なく、WindowsOS全体で動作します。
- ショートカットで入力対象となるExcelのワークブック名、ワークシート名、OK、NG等の入力文字、及び入力場所は自由に設定可能です。

# 使い方
最新のRelease.zipをダウンロードして展開します。<br>
ダウンロードしようとするとブラウザが警告を出すかもですが、それはこのソフトがWindowsOS全体で適用できるショートカットを作るために、キー入力を監視しているからですね。<br>
つまりキーロガーが仕込まれているんじゃないかってブラウザは警告してくれているわけです。<br>
もちろん仕込んでないし、ソースコードは公開しているので安心して使ってください。<br>

Release.zipの中に入っているExToolsForExcelText.exeを実行すれば、バックグラウンドで起動し続け、ショートカットを使えるようになります。<br>
Windowsで警告が出るかもしれませんが、それも多分上と同様の理由ですね。気にしなくていいです。<br>

ショートカットを使うには、対象のExcelファイルが開いている状態である必要があります。<br>
タスクトレイのアイコンをクリックすれば、各種設定が行えます。<br>

入力項目としては、<br>
- ワークブック名に、OKとかNGとか記入するExcelのファイル名
- テスト用シート名に、OKとかNGとか記入するExcelのワークシート名
- エビデンス用シート名に、エビデンスの画像を張り付けるシート名
- テスト通過時テキストに、OKとか〇とか、テストが通過したときに入力する文字列
- テスト不可時テキストに、NGとか✕とか、テストがダメだったときに入力する文字列
- テスト番号のカラムに、テスト番号が入力されている列名(エビデンス画像貼り付け時に使用。無いとエビデンス作成時に落ちる不具合を確認済み)
- テスト内容のカラムに、テスト内容が入力されている列名(エビデンス画像貼り付け時に使用)
- テスト結果のカラムに、OKとかNGとかを入力する列名
- 上余白行数に、テスト結果入力後の自動スクロールした際の上余白行数を指定(これ行と列を逆に表記してますね)
- エビデンス画像サイズに、Excelに貼り付ける際のキャプチャ画像の縮小倍率を指定

という風になっております。
これらを設定して、「適用」ボタンを押せば設定の変更が適用され、次回起動時も同じ設定で起動できます。<br>

ショートカット欄よりも下の項目は開発中なので、見なかったことにしてください。

# ライセンス
MITライセンスです。<br>
そしてまだまだベータ版です。<br>

今公開しているのは最小限の機能で、これからもっと作りこんでいくつもりです。<br>
同じ悩みを抱えているエンジニアの方には、開発に参加してもらえると嬉しいです。<br>
言語はC#、.NET Frameworkです。<br>

- コメントの追加
- バグの指摘や修正
- テストコードの作成
- リファクタリング
- 機能追加
- Readmeの追記

全部歓迎です。

以上です。