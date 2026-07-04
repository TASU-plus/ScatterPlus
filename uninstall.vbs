Option Explicit

Dim fso, shell, excelApp
Dim destFolder, destPath, addInName, addInTitle
Dim configFolder, configPath, configFileName

' --- 設定 ---
addInName = "ScatterPlus.xlam"
addInTitle = "ScatterPlus - 散布図設定ツール" ' ※VBAプロジェクト名やAddInsプロパティに準拠
configFileName = "config.ini"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' 1. パスの確定
destFolder = shell.ExpandEnvironmentStrings("%AppData%\Microsoft\AddIns\")
destPath = fso.BuildPath(destFolder, addInName)
configFolder = shell.ExpandEnvironmentStrings("%AppData%\ScatterPlus")
configPath = fso.BuildPath(configFolder, "\" & configFileName)

' 2. Excelでアドインの登録を解除
On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
excelApp.DisplayAlerts = False
excelApp.Visible = False ' バックグラウンドで実行

' 3. アドインの登録解除（チェックを外す）
excelApp.AddIns(addInTitle).Installed = False

excelApp.Quit
Set excelApp = Nothing
On Error GoTo 0

' 4. アドインファイルの削除
If fso.FileExists(destPath) Then
    On Error Resume Next
    fso.DeleteFile destPath, True
    If Err.Number <> 0 Then
        MsgBox "ファイルの削除に失敗しました。Excelが開いている場合は閉じてから再実行してください。", vbCritical, "エラー"
        WScript.Quit
    End If
    On Error GoTo 0
End If

' 5. 設定ファイルの削除
If fso.FolderExists(configFolder) Then

    if fso.FileExists(configPath) then
        On Error Resume Next
        fso.DeleteFile configPath, True
        On Error GoTo 0
    End If

    On Error Resume Next
    fso.DeleteFolder configFolder, True
    On Error GoTo 0

End If

MsgBox "ScatterPlus のアンインストールが完了しました。", vbInformation, "完了"
