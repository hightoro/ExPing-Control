Import-Module .\Modules\UIAutomation\UIAutomation.dll #UI Automationのpowershell用のラッパ
Add-type -AssemblyName microsoft.VisualBasic #WindowのActive化関数のために呼び出す
Add-Type -AssemblyName System.Windows.Forms #Form描画のため
Add-Type -AssemblyName System.Drawing #Form描画のため
Add-Type -AssemblyName System.Net #Networkの情報取得のため
Add-Type -AssemblyName presentationframework
#Add-Type -AssemblyName System.Threading #ver2.0では使用できない

# -------------------------------------------------------------
# 環境
# -------------------------------------------------------------
#$csv_file = ".\sample3.csv"
[UIAutomation.Preferences]::Highlight = $false
$ErrorActionPreference = "Stop"
$cur_dir = (Convert-Path .)

# -------------------------------------------------------------
# Ping-Controlフォーム
# -------------------------------------------------------------
# フォーム全体の設定
$form2 = New-Object System.Windows.Forms.Form -Property @{
    Text = "実行リスト"
    Size = New-Object System.Drawing.Size(820,520)
    StartPosition = "CenterScreen"
    # フォームを最前面に表示
    #Topmost = $True
}
# 起動ボタンの設定
$StartupButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(40,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "ExPing起動"
}
# Pingボタンの設定
$PingButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(170,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Ping"
    Enabled = $False;
}
# ストップボタンの設定
$StopButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(300,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Stop"
    Enabled = $False;
}
# Tracertボタンの設定
$TraceButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(430,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Tracert"
    Enabled = $False;
}
# クリアボタンの設定
$ClearButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(560,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Clear"
    Enabled = $False;
}
# 終了ボタンの設定
$ShutdownButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(690,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "ExPing終了"
    Enabled = $False;
}
# 項番(ラベル)
$NumberLabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,5)
    Size = New-Object System.Drawing.Size(30,20)
    Text = "項番"
    TextAlign = [ System.Drawing.ContentAlignment]::MiddleLeft
}
# 項番
$NumberTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(50,5)
    Size = New-Object System.Drawing.Size(30,20)
    Text = "1"
}
# データグリッドの設定
$dg1 = New-Object System.windows.forms.DataGridView -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(780,320)
    AutoSize = $True
    AllowUserToAddRows = $False
}
# データテーブルのカラム定義とデータグリッドビューへのバインディング
$dt1 = New-Object System.Data.DataTable
[void]$dt1.Columns.Add("項番", [String])
[void]$dt1.Columns.Add("パス", [String])
[void]$dt1.Columns.Add("保存ファイル名", [String])
[void]$dt1.Columns.Add("プロセスID", [Int32])
[void]$dt1.Columns.Add("状態", [String])
$dg1.DataSource = $dt1

# フォームにアイテムを追加
$form2.Controls.Add($StartupButton)
$form2.Controls.Add($PingButton)
$form2.Controls.Add($StopButton)
$form2.Controls.Add($TraceButton)
$form2.Controls.Add($ClearButton)
$form2.Controls.Add($ShutdownButton)
$form2.Controls.Add($NumberLabel)
$form2.Controls.Add($NumberTextBox)
$form2.Controls.Add($dg1)

# OpenFileDialogクラスのインスタンスを作成
$ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    #FileName = "sample3.csv" # はじめのファイル名を指定する
    # はじめに表示されるフォルダを指定する
    # 指定しない（空の文字列）の時は、現在のディレクトリが表示される
    InitialDirectory = (Convert-Path .)
    # [ファイルの種類]に表示される選択肢を指定する
    # 指定しないとすべてのファイルが表示される
    Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*";
    # [ファイルの種類]ではじめに選択されるものを指定する
    # 2番目の「すべてのファイル」が選択されているようにする
    FilterIndex = 1;
    # タイトルを設定する
    Title = "開くファイルを選択してください";
    # ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
    RestoreDirectory = $True;
    # 存在しないファイルの名前が指定されたとき警告を表示する
    # 使うとなぜかフリーズするので無効化･･･
    CheckFileExists = $False;
    # 存在しないパスが指定されたとき警告を表示する
    CheckPathExists = $True;
    # STAでないとOpenFileDialogが起動しないが、下記のオプションを有効にするとなぜかMTAでも動く。意味わからん
    ShowHelp = $True;
}

# イベント
$StartupButton.Add_Click({
    $StartupButton.Enabled = $false
    StartupExPing
    $PingButton.Enabled = $true
    $TraceButton.Enabled = $true
    $ShutdownButton.Enabled = $true
})
$PingButton.Add_Click({
    $PingButton.Enabled = $false
    $TraceButton.Enabled = $false
    $ClearButton.Enabled = $false
    $ShutdownButton.Enabled = $false
    PopupAutoRunning
    RunPingExPing
    $StopButton.Enabled = $true
})
$StopButton.Add_Click({
    $StopButton.Enabled = $false
    PopupAutoRunning
    StopPingExping
    SavePingExPing
    $PingButton.Enabled = $true
    $TraceButton.Enabled = $true
    $ClearButton.Enabled = $true
    $ShutdownButton.Enabled = $true
})
$TraceButton.Add_Click({
    $PingButton.Enabled = $false
    $TraceButton.Enabled = $false
    $ClearButton.Enabled = $false
    $ShutdownButton.Enabled = $false
    RunTraceExPing
    WaitTraceExPing
    SaveTraceExPing
    $PingButton.Enabled = $true
    $TraceButton.Enabled = $true
    $ClearButton.Enabled = $true
    $ShutdownButton.Enabled = $true
})
$ClearButton.Add_Click({
    $ClearButton.Enabled = $False
    ClearResultExPing
})
$ShutdownButton.Add_Click({
    $ShutdownButton.Enabled = $False
    ShutdownExPing
    $PingButton.Enabled = $False
    $StopButton.Enabled = $False
    $TraceButton.Enabled = $False
    $ClearButton.Enabled = $False
    $StartupButton.Enabled = $True
})

function StartupExPing {
    # ExPingの起動
    $dt1 | ForEach {
        try
        {
            $ps = Start-Process -PassThru -FilePath  ($_.パス+"\ExPing.exe")
            $_.プロセスID = $ps.ID
            #$global:plist += Start-Process -PassThru -FilePath  ($_.Path+"\ExPing.exe")
            $_.状態 = "起動"
        }
        catch [InvalidOperationException]
        {
            $_.状態 = "ファイルが見つかりません"
        }
    }
}
function PopupAutoRunning {
    #警告メッセージの表示
    $m = " マウスから手を放して、Enterを押してください。\n自動操作の間はマウスを操作しないようにしてください"
    [System.Windows.Forms.MessageBox]::Show($m, "", "OK", "Information")
}
function RunPingExPing {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        # ExPingのコントロール取得
        #Write-Host ("プロセスID："+$_.プロセス)
        $win = Get-UiaWindow -ProcessId $_.プロセスID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F5);
        $_.状態 = "Ping実行中"
    }
}
function StopPingExping {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        # ExPingのコントロール取得
        $win = Get-UiaWindow -ProcessId $_.プロセスID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick #| Out-Null #なんかすげえ遅い
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F4); #すげえ早い
        $_.状態 = "Ping停止"
    }
    #Start-Sleep -m 1000;

    # Ping実行が止まるまで待機
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        While($true){
            $win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick | Out-Null
            $str = $win `
            | Get-UiaStatusBar -AutomationId 'StatusBar' -Class 'TStatusBar' `
            | Get-UiaEdit -AutomationId 'StatusBar.Pane0' `
            | Get-UiaEditText
            #Write-host $str
            if($str -eq '終了しました。' ){ break }
            Start-Sleep -m 1000
        }
    }
    #Write-host "End Ping"
}
function SavePingExping {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        #ファイルパス＋ファイル名
        $save_dir1 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.保存ファイル名+"_Ping結果_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"
        $save_dir2 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.保存ファイル名+"_Ping統計_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"
        #Write-Host $save_dir1

        $win = Get-UiaWindow -ProcessId $_.プロセスID;

        # Ping結果の保存ダイアログを開く
        <#
        $win | Get-UiaMenuBar -AutomationId 'MenuBar' -Name 'アプリケーション' `
             | Get-UiaMenuItem -AutomationId 'Item 1' -Name 'ファイル(F)' `
             | Invoke-UIAMenuItemExpand `
             | Get-UiaMenu -Class '#32768' -Name 'ファイル(F)' `
             | Get-UiaMenuItem -AutomationId 'Item 6' -Name 'Ping 結果の保存(K)' `
             | Invoke-UIAMenuItemClick | Out-Null
        #>
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_K);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);

        Start-Sleep -m 1000;
        # 保存するかどうか聞かれた時のダイアログを取得
        $sub = Get-UIAActivewindow
        # ファイル名を入力
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name 'ファイル名(N):' `
        | Set-UIAEditText $save_dir1 | Out-Null
        # 保存ボタンを押す
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '保存(S)' `
        | Invoke-UiaButtonClick | Out-Null

        # Ping統計の保存ダイアログを開く
        <#
        $win | Get-UiaMenuBar -AutomationId 'MenuBar' -Name 'アプリケーション' `
             | Get-UiaMenuItem -AutomationId 'Item 1' -Name 'ファイル(F)' `
             | Invoke-UIAMenuItemExpand `
             | Get-UiaMenu -Class '#32768' -Name 'ファイル(F)' `
             | Get-UiaMenuItem -AutomationId 'Item 7' -Name 'Ping 統計の保存(T)' `
             | Invoke-UIAMenuItemClick | Out-Null
        #>
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_T);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        Start-Sleep -m 1000;
        # 保存するかどうか聞かれた時のダイアログを取得
        $sub = Get-UIAActivewindow
        # ファイル名を入力
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name 'ファイル名(N):' `
        | Set-UIAEditText $save_dir2 | Out-Null
        # 保存ボタンを押す
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '保存(S)' `
        | Invoke-UiaButtonClick | Out-Null
    }
}
function RunTraceExPing {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        # ExPingのコントロール取得
        $win = Get-UiaWindow -ProcessId $_.プロセスID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F6);
    }
    #Start-Sleep -m 1000;
}
function WaitTraceExPing {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        While($true){
            $win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick | Out-Null
            $str = $win `
            | Get-UiaStatusBar -AutomationId 'StatusBar' -Class 'TStatusBar' `
            | Get-UiaEdit -AutomationId 'StatusBar.Pane0' `
            | Get-UiaEditText
            #Write-host $str
            if($str -eq '終了しました。' ){ break }
            Start-Sleep -m 1000
        }
    }
}
function SaveTraceExPing {
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {
        #ファイルパス＋ファイル名
        $save_dir3 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.保存ファイル名+"_Trace結果_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"

        $win = Get-UiaWindow -ProcessId $_.プロセスID;

        # Ping結果の保存ダイアログを開く
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::SHIFT);
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_K);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::SHIFT);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        Start-Sleep -m 1000;
        # 保存するかどうか聞かれた時のダイアログを取得
        $sub = Get-UIAActivewindow
        # ファイル名を入力
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name 'ファイル名(N):' `
        | Set-UIAEditText $save_dir3 | Out-Null
        # 保存ボタンを押す
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '保存(S)' `
        | Invoke-UiaButtonClick | Out-Null
    }
}
function ClearResultExPing {
    #結果をクリアする
    $dt1 | Where {$_.プロセスID -ne [System.DBNULL]::Value} | ForEach {

        $win = Get-UiaWindow -ProcessId $_.プロセスID;

        # Ping結果の保存ダイアログを開く
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::DELETE);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
    }
}
function ShutdownExPing {
    $dt1 `
    | Where {$_.プロセスID -ne [System.DBNULL]::Value} `
    | Where {(Get-Process -id $_.プロセスID -ErrorAction SilentlyContinue) -ne $null} `
    | ForEach {
        Stop-Process -id $_.プロセスID;
        $_.プロセスID = [System.DBNULL]::Value
        $_.状態 = "未起動"
    }
}
function read_csv {
    try
    {
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            Get-Content $ofd.FileName -ErrorAction Stop | ConvertFrom-Csv | ForEach {
                [void]$dt1.Rows.Add($_.No,$_.パス,$_.保存ファイル名,$null,$null)
            }
        } else {
            exit
        }
    }
    catch [System.Management.Automation.ActionPreferenceStopException]
    {
        $str = $_.CategoryInfo.TargetName + "`nが見つかりません"
        [void][System.Windows.Forms.MessageBox]::Show($str,"エラー","OK","Warning" )
        read_csv
    }
}


read_csv
$form2.ShowDialog()
