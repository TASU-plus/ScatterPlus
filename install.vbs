Option Explicit

Dim fso, shell, excelApp, targetAddIn
Dim srcPath, destFolder, destPath
Dim addInName

' --- 設定 ---
addInName = "ScatterPlus.xlam"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' 1. パスの確定
' スクリプトと同じフォルダにあるアドインファイルを指定
srcPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), addInName)
destFolder = shell.ExpandEnvironmentStrings("%AppData%\Microsoft\AddIns\")
destPath = fso.BuildPath(destFolder, addInName)

' 2. アドインフォルダの存在確認と作成
If Not fso.FolderExists(destFolder) Then
    On Error Resume Next
    fso.CreateFolder(destFolder)
    If Err.Number <> 0 Then
        MsgBox "アドインフォルダの作成に失敗しました。権限を確認してください。", vbCritical, "エラー"
        WScript.Quit
    End If
    On Error GoTo 0
End If

' 3. ファイルの物理コピー
On Error Resume Next
fso.CopyFile srcPath, destPath, True

If Err.Number <> 0 Then
    ' コピー失敗時はExcelが掴んでいると判断
    MsgBox "ファイルのコピーに失敗しました。" & vbCrLf & vbCrLf & _
           "Excelが起動している場合は、すべてのExcelウィンドウを閉じてから再度実行してください。", _
           vbExclamation, "Excelを閉じてください"
    WScript.Quit
End If
On Error GoTo 0

' 4. Excelでの有効化処理
Set excelApp = CreateObject("Excel.Application")
excelApp.DisplayAlerts = False
excelApp.Visible = False ' バックグラウンドで実行

' 【重要】Addを成功させるための「空ブックの追加」
excelApp.Workbooks.Add

' 5. アドインの登録 (Add)
On Error Resume Next
Set targetAddIn = excelApp.AddIns.Add(destPath)

If Err.Number <> 0 Or targetAddIn Is Nothing Then
    ' Addに失敗した場合
    excelApp.Quit
    MsgBox "Excelへのアドイン登録（Add）に失敗しました。" & vbCrLf & _
           "お手数ですが、Excelの「オプション」＞「アドイン」から手動でインストールを行ってください。", _
           vbCritical, "手動インストールのお願い"
    WScript.Quit
End If
On Error GoTo 0

' 6. 有効化 (Installed)
On Error Resume Next
targetAddIn.Installed = True
On Error GoTo 0

' 7. 終了処理
excelApp.Quit
Set excelApp = Nothing

MsgBox "ScatterPlus のセットアップが正常に完了しました！" & vbCrLf & _
       "Excelを起動してリボンにタブが表示されているか確認してください。", _
       vbInformation, "インストール完了"
