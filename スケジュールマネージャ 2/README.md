# Outlook 共通空き時間検索

Windows の**従来版 Outlook**にサインインしている利用者が、自分と入力した最大10人の共通空き時間を調べる画面アプリです。Microsoft 365 のアプリ登録や Graph API は使用しません。Python の追加パッケージも不要です。Python 標準ライブラリの `tkinter` と `subprocess` から、Windows 標準の PowerShell を介して Outlook の COM 機能を呼び出します。

## 必要な環境

- Windows、Python 3.10 以降（`tkinter` / Tcl/Tk を含む標準インストール）、Windows PowerShell 5.1
- **従来版 Outlook for Windows** がインストールされ、メールアカウントにサインインしていること。検索前に Outlook を開いてください。
- 対象者の空き時間を、その Outlook 利用者が参照できること。参照可否は組織の予定表共有設定に従います。

「新しい Outlook」や Web 版 Outlook では、この方法は利用できません。Windows と Outlook の表示時刻が一致している環境を想定しています。

## 起動と使い方

`outlook_free_time.py` と `outlook_freebusy.ps1` を同じフォルダーに置き、PowerShell から次を実行します。

```powershell
python outlook_free_time.py
```

PowerShell では `cd 'C:\Users\MG\OneDrive\ドキュメント\ChatGPT\スケジュールマネージャ 2'` でこのフォルダーに移動してもかまいません。`pythonw.exe .\outlook_free_time.py` でも起動できますが、PowerShell 自体の画面は残ります。ターミナルを開かずに起動する場合は、同じフォルダーの `outlook_free_time.pyw` をダブルクリックしてください（`.pyw` が Python に関連付けられている環境）。

画面の10個の欄に相手のメールアドレスを1件ずつ入力します。自分は Outlook の現在の利用者から自動取得するため、入力不要です。空欄は無視され、重複は1人として扱います。開始日・終了日を `YYYY-MM-DD` で入力すると両日を含めて検索します。時間は30分刻みで、初期値は09:00～18:00です。1日全体なら00:00～24:00に変更してください。検索期間は最大31日です。

検索結果は、自分と入力した全員の予定が「空き」である30分枠を連続した時間帯にまとめて表示します。仮予定・予定あり・外出中・他の場所で作業中は空き扱いしません。自分または相手の予定取得に失敗した場合は、空き時間を推測せずエラーを表示します。過去の時間帯は表示しません。複数の Outlook アカウントを設定している場合、「自分」は Outlook の現在の利用者を指します。

Microsoft の資料: [Recipient.FreeBusy](https://learn.microsoft.com/en-us/office/vba/api/Outlook.Recipient.FreeBusy)、[新しい Outlook の COM 対応状況](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/overview-new-outlook-windows)
