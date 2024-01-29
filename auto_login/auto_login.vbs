Dim objWshShell
Dim GC

'アクセスしたいURLを設定
const url = "" 

'ログインユーザIDを設定
const user_id = ""

'ログインパスワードを設定
const pass = ""

Set objWshShell = WScript.CreateObject("WScript.Shell")
Set GC = CreateObject("WScript.Chell")

'シークレットモードで起動
GC.Run("chrome.exe --incognito -url " & url)
objWshShell.AppActive "chorome.exe"
'4000ミリ秒待機
WScript.Sleep 4000
'全画面設定
objWshShell.SendKeys "{F11}"

'ログインユーザIDをクリップボードに保存
Call CreateObject("Wscript.Shell").Run("%COMSPEC% /C echo " & user_id & "|clip" , 0) 
WScript.Sleep 300
'貼り付け
objWshShell.SendKeys("^V")

'TABで移動
objWshShell.SendKeys "{Tab}"
WScript.Sleep 300

'ログインパスワードをクリップボードに保存
Call CreateObject("Wscript.Shell").Run("%COMSPEC% /C echo " & pass & "|clip" , 0) 
WScript.Sleep 300
'貼り付け
objWshShell.SendKeys("^V")
'エンターキーでSUBMIT
objWshShell.SendKeys "{Enter}"
WScript.Sleep 300

'全画面終了
objWshShell.SendKeys "{F11}"




