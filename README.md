# Office アドイン タスクペイン React JS

このプロジェクトは、React で構築された Outlook アドインです。Outlook 内でタスクペインのインターフェースを提供します。

-----

## インストール

1.  **リポジトリをクローンする:**
    ```bash
    git clone https://github.com/parvezamm3/mailai_addin.git
    ```
2.  **プロジェクトディレクトリに移動する:**
    ```bash
    cd mailai_addin
    ```
3.  **依存関係をインストールする:**
    ```bash
    npm install
    ```

-----

## 開始方法

1.  **開発サーバーを起動し、アドインをサイドロードする:**

    ```bash
    npm start
    ```

    このコマンドを実行すると、開発サーバーが起動し、アドインがサイドロードされた状態で Outlook が自動的に開きます。

2.  **Outlook でアドインを開く:**

      * Outlook で、メールを開きます。
      * 「アプリ」ボタンをクリックします。
      * 「MailAI」を選択して、アドインを開きます。


## 承認
Flask バックエンドに、メールを自動的に読み込むための Graph API 権限を承認するには、この URL を使用してください。
    ([https://equipped-externally-stud.ngrok-free.app/outlook-oauth2callback](https://equipped-externally-stud.ngrok-free.app/outlook-oauth2callback))