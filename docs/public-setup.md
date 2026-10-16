# セットアップと更新手順

## 閲覧のみの場合

GitHubのREADMEとgasフォルダを閲覧できます。Google認証や設定変更は不要です。

## 自分の環境へ導入する場合

1. リポジトリをcloneし、独立したApps Scriptプロジェクトを作成します。
2. `gas/00_Config.example.js`を`gas/00_Config.js`へコピーし、自分のDrive・フォーム・作業スプレッドシート・通知先等を設定します。作業スプレッドシートの内容は処理時に置き換わります。専用の空ファイルを用意してください。
3. `.clasp.example.json`を`.clasp.json`へコピーし、自分のGASプロジェクトのscriptIdを指定します。rootDirはgasです。
4. claspを公式npmパッケージからインストールし、`clasp login`で自分のGoogleアカウントを認証します。
5. `clasp status`で対象を確認してから`clasp push`します。設定見本は`.claspignore`により同期対象外です。
6. GAS側でDrive API・Sheets APIを有効にし、マニフェストに記載された権限を確認して認可します。
7. メール自動収集・ファイル提出3項目・対象シート名と開始列／終了列をフォームに用意します。質問タイトルは`05_FormIntake.js`のrequired配列と合わせます。
8. `setupIndependentDataForm`を手動実行し、開始行・出力終了行・ヘッダー有無・文字コード・改行・BOMの必須11項目を追加します。同じ関数を再実行しても質問は重複しません。
9. `setupProcessingQueue`を手動実行し、フォーム送信トリガーと1分ごとの処理トリガーを設定します。既存環境で設定済みの場合、トリガーの再設定は不要です。

デモファイルのハッシュによる限定はありません。ユーザーそれぞれの3ファイルとフォーム回答をもとに処理します。開始行はヘッダーを含めるかどうかも考慮して明示し、出力範囲の行数・列数・内容が提出する正解の後CSV全体に一致するよう指定してください。UTF-8またはShift_JISを選択でき、Shift_JISではBOMを「なし」にします。

既存環境を更新する場合は、コード同期後に`setupIndependentDataForm`を1回実行してください。追加条件が欠けた古いフォーム回答はデモ値で補完せず、修正・再提出の案内を送ります。完了済み案件は再実行しません。

メール通知には実際のフォーム回答IDと自動収集メールアドレスの一致が必要です。フォームを介さない手動実行は、`examples/job.example.json`を参考に作成したジョブJSONのファイルIDをJOB_FILE_IDに設定し、`runConfiguredJob`で実行できます。

## オーナーの通常更新

1. コードを修正し、`node --test tests/*.test.cjs`を実行します。
2. gitの差分と登録対象を確認してcommit・GitHubへpushします。
3. GitHub Actionsのチェック結果を確認します。
4. `clasp status`で対象を確認し、`clasp push`で専用GASへ同期します。
5. フォームでテストし、GAS実行履歴・照合結果・受信メールを確認します。

GitHub Actionsはテストだけを行います。Google認証情報をGitHubに登録する自動デプロイは設定していません。

## 登録対象外

`gas/00_Config.js`（実環境の接続設定）、`.clasp.json`、`.clasprc.json`、`.env`、実環境のIDを含む`gas/08_DemoTest.js`、提出資料、生成物、作業中の非公開記録は`.gitignore`で除外します。認証情報や顧客データをcommitしないでください。
