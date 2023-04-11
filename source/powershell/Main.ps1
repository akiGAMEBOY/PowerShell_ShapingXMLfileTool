#################################################################################
# 処理名　｜ShapingXMLfileTool（メイン処理）
# 機能　　｜XMLファイルを整形するツール
#--------------------------------------------------------------------------------
# 戻り値　｜下記の通り。
# 　　　　｜   0: 正常終了
# 　　　　｜-001: エラー 設定ファイル読み込み
# 　　　　｜-211: エラー 参照できないファイル
# 　　　　｜-201: エラー メイン - 設定ファイル読み込み
# 　　　　｜-311: エラー 取り込んだデータが0件
# 　　　　｜-401: エラー XMLファイル内にバージョン情報なし
# 　　　　｜-411: エラー 対象のタグがない
# 　　　　｜-421: エラー v1.0の必須項目エラー
# 　　　　｜-422: エラー v2.0の必須項目エラー
# 　　　　｜-423: エラー 既定外のバージョンのXMLファイル
# 　　　　｜-431: エラー コピーバックアップエラー
# 　　　　｜-501: エラー XMLファイルの置換失敗
# 　　　　｜-502: エラー コピーバックアップ - フォルダのコピー
# 　　　　｜-901: エラー メイン - 処理中断
# 引数　　｜-
#################################################################################
# 設定
# 定義されていない変数があった場合にエラーとする
Set-StrictMode -Version Latest
# アセンブリ読み込み（フォーム用）
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# try-catchの際、例外時にcatchの処理を実行する
$ErrorActionPreference = "Stop"
# 定数
[System.String]$c_config_file = "setup.ini"
## 構造体
[PSCustomObject]$c_systemver_10 = @{
    SystemVersion = "SystemVersion=`"1`.0`""
    SortOrder = @(
        " SystemVersion=`".*?`"",
        " PurchaseOrderNumber=`".*?`"",
        " OrderDate=`".*?`"",
        " Remarks=`".*?`""
    )
    Required = @(
        '1',
        '1',
        '1',
        '1'
    )
}
[PSCustomObject]$c_systemver_20 = @{
    SystemVersion = "SystemVersion=`"2`.0`""
    SortOrder = @(
        " SystemVersion=`".*?`"",
        " PurchaseOrderNumber=`".*?`"",
        " OrderDate=`".*?`"",
        " CreateDate=`".*?`"",
        " Remarks=`".*?`""
    )
    Required = @(
        '1',
        '1',
        '1',
        '0',
        '1'
    )
}
# DEBUG
[System.Boolean]$c_debug = $true
# Function
#################################################################################
# 処理名　｜ExpandString
# 機能　　｜文字列を展開（先頭桁と最終桁にあるダブルクォーテーションを削除）
#--------------------------------------------------------------------------------
# 戻り値　｜String（展開後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function ExpandString {
    param ([System.String]$target_str)
    [System.String]$expand_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq "`"") -and
                ($target_str.Substring($target_str.Length - 1, 1) -eq "`"")) {
            # ダブルクォーテーション削除
            $expand_str = $target_str.Substring(1, $target_str.Length - 2)
        }
    }

    return $expand_str
}

#################################################################################
# 処理名　｜ConfirmYesno_winform
# 機能　　｜YesNo入力（Windowsフォーム）
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno_winform {
    param (
        [System.String]$prompt_message
    )
    [System.Boolean]$return = $false

    # フォームの作成
    [System.Windows.Forms.Form]$form = New-Object System.Windows.Forms.Form
    $form.Text = "実行前の確認"
    $form.Size = New-Object System.Drawing.Size(460,210)
    $form.StartPosition = "CenterScreen"
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("${root_dir}\source\icon\shell32-296.ico")
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    # ピクチャボックス作成
    [System.Windows.Forms.PictureBox]$pic = New-Object System.Windows.Forms.PictureBox
    $pic.Size = New-Object System.Drawing.Size(32, 32)
    $pic.Image = [System.Drawing.Image]::FromFile("${root_dir}\source\icon\shell32-296.ico")
    $pic.Location = New-Object System.Drawing.Point(30,30)
    $pic.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    # ラベル作成
    [System.Windows.Forms.Label]$label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(85,30)
    $label.Size = New-Object System.Drawing.Size(350,80)
    $label.Text = $prompt_message
    $font = New-Object System.Drawing.Font("ＭＳ ゴシック",12)
    $label.Font = $font
    # OKボタンの作成
    [System.Windows.Forms.Button]$btnOkay = New-Object System.Windows.Forms.Button
    $btnOkay.Location = New-Object System.Drawing.Point(255,120)
    $btnOkay.Size = New-Object System.Drawing.Size(75,30)
    $btnOkay.Text = "OK"
    $btnOkay.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # Cancelボタンの作成
    [System.Windows.Forms.Button]$btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(345,120)
    $btnCancel.Size = New-Object System.Drawing.Size(75,30)
    $btnCancel.Text = "キャンセル"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    # ボタンの紐づけ
    $form.AcceptButton = $btnOkay
    $form.CancelButton = $btnCancel
    # フォームに紐づけ
    $form.Controls.Add($pic)
    $form.Controls.Add($label)
    $form.Controls.Add($btnOkay)
    $form.Controls.Add($btnCancel)
    # フォーム表示
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $return = $true
    } else {
        $return = $false
    }
    $pic.Image.Dispose()
    $pic.Image = $null
    $form = $null

    return $return
}

#################################################################################
# 処理名　｜RetrieveMessage
# 機能　　｜メッセージ内容を取得
#--------------------------------------------------------------------------------
# 戻り値　｜String（メッセージ内容）
# 引数　　｜target_code; 対象メッセージコード, append_message: 追加メッセージ（任意）
#################################################################################
Function RetrieveMessage {
    param (
        [System.String]$target_code,
        [System.String]$append_message=''
    )
    [System.String]$return = ''

    [System.String[][]]$messages = @(
        ("0", "正常終了"),
        ("-1", "設定ファイルの読み込みでエラー。"),
        ("-111", "必須項目が未入力。"),
        ("-112", "数値項目で数値以外が入力。"),
        ("-113", "二重登録あり（1行内に複数の階層を入力）。"),
        ("-114", "矛盾あり（1行目が0階層以外で設定）。"),
        ("-115", "前後の階層関係に誤り。"),
        ("-211", "参照できないファイルがあり。"),
        ("-212", "参照できないフォルダがあり。"),
        ("-311", "取り込んだデータが0件。"),
        ("-401", "XMLファイル内にバージョン情報がない。"),
        ("-411", "整形対象のタグがない。"),
        ("-421", "SystemVersion1.0の必須項目チェックでエラー。"),
        ("-422", "SystemVersion2.0の必須項目チェックでエラー。"),
        ("-423", "既定バージョン以外のXMLファイル。"),
        ("-431", "既定バージョン以外のXMLファイル。"),
        ("-501", "XMLファイルの置換で失敗。"),
        ("-901", "処理をキャンセル。"),
        ("-999", "例外が発生。")
    )

    for ([System.Int32]$i = 0; $i -lt $messages.Length; $i++) {
        if ($messages[$i][0] -eq $target_code) {
            $sbtemp=New-Object System.Text.StringBuilder
            @("$($messages[$i][1])`r`n",`
              "${append_message}`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $return = $sbtemp.ToString()
            break
        }
    }
    
    return $return
}

#################################################################################
# 処理名　｜IsExistsAttribute
# 機能　　｜属性の入力チェック
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜target_path: 対象XMLファイル, target_ver: 対象バージョン
#################################################################################
function IsExistsAttribute {
    param (
        [System.String]$target_path,
        [PSCustomObject]$target_ver
    )
    [System.Boolean]$return = $true
    [System.String]$match_str = ''

    for ([System.Int32]$i = 0; $i -lt $target_ver.SortOrder.Length; $i++) {
        $match_str = [Regex]::Matches((Get-Content $target_path), $target_ver.SortOrder[$i]) | ForEach-Object {$_.Value}

        if (($null -eq $match_str) -And `
            ($target_ver.Required[$i] -eq '1')) {
                $return = $false
                break
        }
    }

    return $return
}

#################################################################################
# 処理名　｜SortorderAttribute
# 機能　　｜属性の並び替え
#--------------------------------------------------------------------------------
# 戻り値　｜String（並び替え後のタグ）
# 引数　　｜target_path: 対象XMLファイル, target_ver: 対象バージョン
#################################################################################
function SortorderAttribute {
    param (
        [System.String]$target_path,
        [PSCustomObject]$target_ver
    )
    [System.String]$return = '<PurchaseOrder'
    [System.String]$match_str = ''

    # 対象文字列を読み込み
    [System.String]$target_str = [Regex]::Matches((Get-Content -Raw $target_path), "<PurchaseOrder[\s\S]*?>") | ForEach-Object {$_.Value}

    # 並び替え
    foreach($item in $target_ver.SortOrder) {
        $match_str = [Regex]::Matches($target_str, $item) | ForEach-Object {$_.Value}
        $sbtemp=New-Object System.Text.StringBuilder
        @($return,`
          $match_str)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $return = $sbtemp.ToString()
    }
    $sbtemp=New-Object System.Text.StringBuilder
    @($return,`
      ">")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $return = $sbtemp.ToString()

    return $return
}

#################################################################################
# 処理名　｜RemoveUnnecessaryparts
# 機能　　｜不要な一部属性の削除
#--------------------------------------------------------------------------------
# 戻り値　｜String（削除後のタグ）
# 引数　　｜target_str: 対象文字列
#################################################################################
function RemoveUnnecessaryparts {
    param (
        [System.String]$target_str
    )
    [System.String]$return = $target_str
    [System.String]$match_str = [Regex]::Matches($target_str," Remarks=`".*?`"") | ForEach-Object {$_.Value}

    if ($null -eq $match_str) {
        return $match_str
    }

    $return = [Regex]::Replace($target_str, $match_str, "", "IgnoreCase")

    return $return
}

#################################################################################
# 処理名　｜DiffTextfile
# 機能　　｜テキスト形式のファイルを比較
#--------------------------------------------------------------------------------
# 戻り値　｜-
# 引数　　｜fromfile：比較元ファイル、tofile：比較先ファイル
# 　　　　　(任意)full：-Full指定で全ての差異表示
#################################################################################
function DiffTextfile {
    param (
        [System.String]$fromfile,
        [System.String]$tofile,
        [Switch]$full
    )
    [System.Int32]$maxrow = 40
    [System.Int32]$rowcount = 0
  
    # ウィンドウサイズの変更
    If (-Not $c_debug) {
        [System.Management.Automation.Host.PSHostRawUserInterface]$userinterface = $host.UI.RawUI
        [System.ValueType]$windowsize = $userinterface.WindowSize
        $userinterface.WindowSize = New-Object System.Management.Automation.Host.Size(120,43)
    }
  
    # 比較処理
    [System.String]$line = ""
    [System.String]$forecolor = ""
    Compare-Object (Get-Content $fromfile) (Get-Content $tofile) -IncludeEqual:$full |
        ForEach-Object {
        if ($_.SideIndicator -eq "=>")
        {
            # 修正後に存在する行（追加または変更された行）
            $line = "[ + ] " + $_.InputObject
            $forecolor = "Red"
        } elseif ($_.SideIndicator -eq "<=") {
            # 修正後に存在しない行（削除または変更された行）
            $line = "[ - ] " + $_.InputObject
            $forecolor = "DarkGray"
        } elseif ($full) {
            # 変更がない行
            $line = "[ = ] " + $_.InputObject
            $forecolor = "White"
        }
        Write-Host $line -ForegroundColor $forecolor
        $rowcount++
        # 最大行数まで達した場合、画面を一時停止
        if ($rowcount -ge $maxrow) {
            $rowcount = 0
            Write-Host ''
            Read-Host ' --- 次のページへ [ Enter ] / 中断 [ Ctrl + C ] --- '
        }
    }
    # ウィンドウサイズの戻し
    Write-Host ''
    Write-Host '--- 比較終了 [ Enter ] ---'
    Read-Host | Out-Null
    if (-Not $c_debug) {
        $userinterface.WindowSize = $windowsize
    }
}

#################################################################################
# 処理名　｜ReplaceXmlfile
# 機能　　｜XMLファイルの置換処理
#--------------------------------------------------------------------------------
# 戻り値　｜Int（0：成功, -501：失敗）
# 引数　　｜target_path；対象XMLファイル, target_str: 対象文字列
#################################################################################
function ReplaceXmlfile {
    param (
        [System.String]$target_path,
        [System.String]$target_str
    )
    [System.Int32]$result = 0
    [System.String]$xmldata = [System.IO.File]::ReadAllText($target_path)
    $xmldata = [Regex]::Replace($xmldata, "<PurchaseOrder[\s\S]*?>", $afterdel)

    try {
        [System.IO.File]::WriteAllText($target_path, $xmldata)
        # DEBUG
        if ($c_debug) {
            $tag_after = [Regex]::Matches((Get-Content $target_path),"<PurchaseOrder[\s\S]*?>") | ForEach-Object {$_.Value}
            $sbtemp=New-Object System.Text.StringBuilder
            @("DEBUG : 下記内容で置換しました。`r`n", `
              "`r`n",`
              "--[置換後]------------------------------------------------`r`n", `
              "${tag_after}`r`n",`
              "----------------------------------------------------------`r`n",`
              "`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkYellow
        }
    } catch {
        $result = -501
    }

    return $result
}

#################################################################################
# 処理名　｜ShapingXMLfile
# 機能　　｜XMLファイル整形処理
#--------------------------------------------------------------------------------
# 戻り値　｜Int
# 　　　　｜   0: 正常終了
# 　　　　｜-311: エラー 取り込んだデータが0件
# 　　　　｜-401: エラー XMLファイル内にバージョン情報なし
# 　　　　｜-411: エラー 対象のタグがない
# 　　　　｜-421: エラー v1.0の必須項目エラー
# 　　　　｜-422: エラー v2.0の必須項目エラー
# 　　　　｜-423: エラー 既定外のバージョンのXMLファイル
# 　　　　｜-431: エラー コピーバックアップエラー
# 引数　　｜target_path; 対象XMLファイル
#################################################################################
Function ShapingXMLfile {
    param (
        [System.String]$target_path
    )
    [System.Int32]$result = 0
    [System.String]$prompt_message = ''

    # データ件数チェック
    $xmldata = (Get-Content $target_path)
    if ($null -eq $xmldata) {
        $result = -311
    }

    # 属性バージョンの存在チェック
    if ($result -eq 0) {
        $xml_systemver = [Regex]::Matches($xmldata, "SystemVersion=`".*?`"") | ForEach-Object {$_.Value}
        if ($null -ne $xml_systemver) {
            $sbtemp=New-Object System.Text.StringBuilder
            @("確認　: バージョンの確認`r`n",`
              "　　　　バージョン[${xml_systemver}]`r`n",`
              "`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkYellow
        } else {
            $result = -401
        }
    }

    # XMLのタグ／属性の値チェック
    if ($result -eq 0) {
        # タグの必須チェック
        $target_tag = [Regex]::Matches((Get-Content -Raw $target_path), "<PurchaseOrder[\s\S]*?>") | ForEach-Object {$_.Value}

        if ($null -eq $target_tag) {
            $result = -411
        }

        # 属性の必須チェック
        if ($result -eq 0) {
            if ($xml_systemver -eq $c_systemver_10.SystemVersion) {
                if (-Not(IsExistsAttribute $target_path $c_systemver_10)) {
                    $result = -421
                }
            } elseif ($xml_systemver -eq $c_systemver_20.SystemVersion) {
                if (-Not(IsExistsAttribute $target_path $c_systemver_20)) {
                    $result = -422
                }
            } else {
                $result = -423
            }
        }
    }

    # 整形（加工・置換）のメイン処理
    ## 置換前のコピーバックアップ
    if ($result -eq 0) {
        # バックアップファイル名やパスを作成
        [System.String]$date_str = Get-Date -Format "yyyyMMdd-HHmmss"
        $sbtemp=New-Object System.Text.StringBuilder
        @($target_path,`
          "_bk",`
          "$date_str")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        [System.String]$backup_path = $sbtemp.ToString()
        # コピーバックアップの実行
        try {
            Copy-Item $target_path -Recurse $backup_path
        } catch {
            $result = -431
        }
    }

    ## 整形（加工と置換）
    if ($result -eq 0) {
        # 属性の並び替え
        [System.String]$aftersorting = ''
        if ($xml_systemver -eq $c_systemver_10.SystemVersion) {
            $aftersorting = SortorderAttribute $target_path $c_systemver_10
        } elseif ($xml_systemver -eq $c_systemver_20.SystemVersion) {
            $aftersorting = SortorderAttribute $target_path $c_systemver_20
        }
        # 不要な一部属性の削除
        [System.String]$afterdel = RemoveUnnecessaryparts $aftersorting
        
        # XMLファイルの置換処理
        $result = ReplaceXmlfile $target_path $afterdel
        # DEBUG
        if ($c_debug) {
            DiffTextfile $backup_path $target_path
        }
    }

    return $result
}

#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
[System.Int32]$result = 0
[System.String]$prompt_message = ''
[System.String]$result_message = ''
[System.String]$append_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 初期設定
## ディレクトリの取得
[System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent
Set-Location $current_dir"\..\.."
[System.String]$root_dir = (Convert-Path .)
## 設定ファイル読み込み
$sbtemp=New-Object System.Text.StringBuilder
@("$current_dir",`
"\",`
"$c_config_file")|
ForEach-Object{[void]$sbtemp.Append($_)}
[System.String]$config_fullpath = $sbtemp.ToString()
try {
    [System.Collections.Hashtable]$param = Get-Content $config_fullpath -Raw -Encoding UTF8 | ConvertFrom-StringData
    # 対象ファイル
    [System.String]$Targetfile=ExpandString($param.Targetfile)

    $sbtemp=New-Object System.Text.StringBuilder
    @("通知　　　: 設定ファイル読み込み`r`n",`
    "　　　　　　設定ファイルの読み込みが正常終了しました。`r`n",`
    "　　　　　　対象: [${config_fullpath}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    Write-Host $prompt_message
}
catch {
    $result = -1
    $append_message = "　　　　　　エラー内容: [${config_fullpath}$($_.Exception.Message)]`r`n"
    $result_message = RetrieveMessage $result $append_message
}
## 対象ファイルの存在チェック
if ($result -eq 0) {
    [System.String]$target_path = "${root_dir}\${Targetfile}"
    if (-Not(Test-Path $target_path)) {
        $result = -211
        $append_message = "　　　　　　対象: [${target_path}]`r`n"
        $result_message = RetrieveMessage $result $append_message
    }
}

# 実行前のポップアップ
if ($result -eq 0) {
    $sbtemp=New-Object System.Text.StringBuilder
    # 実行有無の確認
    @("XMLファイルの整形処理を実行します。`r`n",`
      "処理を続行しますか？`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    If (ConfirmYesno_winform $prompt_message) {
        $result = ShapingXMLfile $target_path
    } else {
        $result = -901
    }
    if ($result -ne 0) {
        $result_message = RetrieveMessage $result
    }
}

# 処理結果の表示
$sbtemp=New-Object System.Text.StringBuilder
if ($result -eq 0) {
    @("処理結果　: 正常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message
}
else {
    @("処理結果　: 異常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n",`
      "　　　　　　",`
      $result_message)|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message -ForegroundColor DarkRed
}

# 終了
exit $result
