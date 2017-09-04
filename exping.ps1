Import-Module .\Modules\UIAutomation\UIAutomation.dll #UI Automation��powershell�p�̃��b�p
Add-type -AssemblyName microsoft.VisualBasic #Window��Active���֐��̂��߂ɌĂяo��
Add-Type -AssemblyName System.Windows.Forms #Form�`��̂���
Add-Type -AssemblyName System.Drawing #Form�`��̂���
Add-Type -AssemblyName System.Net #Network�̏��擾�̂���
Add-Type -AssemblyName presentationframework
#Add-Type -AssemblyName System.Threading #ver2.0�ł͎g�p�ł��Ȃ�

# -------------------------------------------------------------
# ��
# -------------------------------------------------------------
#$csv_file = ".\sample3.csv"
[UIAutomation.Preferences]::Highlight = $false
$ErrorActionPreference = "Stop"
$cur_dir = (Convert-Path .)

# -------------------------------------------------------------
# Ping-Control�t�H�[��
# -------------------------------------------------------------
# �t�H�[���S�̂̐ݒ�
$form2 = New-Object System.Windows.Forms.Form -Property @{
    Text = "���s���X�g"
    Size = New-Object System.Drawing.Size(820,520)
    StartPosition = "CenterScreen"
    # �t�H�[�����őO�ʂɕ\��
    #Topmost = $True
}
# �N���{�^���̐ݒ�
$StartupButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(40,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "ExPing�N��"
}
# Ping�{�^���̐ݒ�
$PingButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(170,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Ping"
    Enabled = $False;
}
# �X�g�b�v�{�^���̐ݒ�
$StopButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(300,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Stop"
    Enabled = $False;
}
# Tracert�{�^���̐ݒ�
$TraceButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(430,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Tracert"
    Enabled = $False;
}
# �N���A�{�^���̐ݒ�
$ClearButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(560,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "Clear"
    Enabled = $False;
}
# �I���{�^���̐ݒ�
$ShutdownButton = New-Object System.Windows.Forms.Button -Property @{
    Location = New-Object System.Drawing.Point(690,360)
    Size = New-Object System.Drawing.Size(90,60)
    Text = "ExPing�I��"
    Enabled = $False;
}
# ����(���x��)
$NumberLabel = New-Object System.Windows.Forms.Label -Property @{
    Location = New-Object System.Drawing.Point(10,5)
    Size = New-Object System.Drawing.Size(30,20)
    Text = "����"
    TextAlign = [ System.Drawing.ContentAlignment]::MiddleLeft
}
# ����
$NumberTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(50,5)
    Size = New-Object System.Drawing.Size(30,20)
    Text = "1"
}
# �f�[�^�O���b�h�̐ݒ�
$dg1 = New-Object System.windows.forms.DataGridView -Property @{
    Location = New-Object System.Drawing.Point(10,30)
    Size = New-Object System.Drawing.Size(780,320)
    AutoSize = $True
    AllowUserToAddRows = $False
}
# �f�[�^�e�[�u���̃J������`�ƃf�[�^�O���b�h�r���[�ւ̃o�C���f�B���O
$dt1 = New-Object System.Data.DataTable
[void]$dt1.Columns.Add("����", [String])
[void]$dt1.Columns.Add("�p�X", [String])
[void]$dt1.Columns.Add("�ۑ��t�@�C����", [String])
[void]$dt1.Columns.Add("�v���Z�XID", [Int32])
[void]$dt1.Columns.Add("���", [String])
$dg1.DataSource = $dt1

# �t�H�[���ɃA�C�e����ǉ�
$form2.Controls.Add($StartupButton)
$form2.Controls.Add($PingButton)
$form2.Controls.Add($StopButton)
$form2.Controls.Add($TraceButton)
$form2.Controls.Add($ClearButton)
$form2.Controls.Add($ShutdownButton)
$form2.Controls.Add($NumberLabel)
$form2.Controls.Add($NumberTextBox)
$form2.Controls.Add($dg1)

# OpenFileDialog�N���X�̃C���X�^���X���쐬
$ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    #FileName = "sample3.csv" # �͂��߂̃t�@�C�������w�肷��
    # �͂��߂ɕ\�������t�H���_���w�肷��
    # �w�肵�Ȃ��i��̕�����j�̎��́A���݂̃f�B���N�g�����\�������
    InitialDirectory = (Convert-Path .)
    # [�t�@�C���̎��]�ɕ\�������I�������w�肷��
    # �w�肵�Ȃ��Ƃ��ׂẴt�@�C�����\�������
    Filter = "CSV�t�@�C��(*.csv)|*.csv|���ׂẴt�@�C��(*.*)|*.*";
    # [�t�@�C���̎��]�ł͂��߂ɑI���������̂��w�肷��
    # 2�Ԗڂ́u���ׂẴt�@�C���v���I������Ă���悤�ɂ���
    FilterIndex = 1;
    # �^�C�g����ݒ肷��
    Title = "�J���t�@�C����I�����Ă�������";
    # �_�C�A���O�{�b�N�X�����O�Ɍ��݂̃f�B���N�g���𕜌�����悤�ɂ���
    RestoreDirectory = $True;
    # ���݂��Ȃ��t�@�C���̖��O���w�肳�ꂽ�Ƃ��x����\������
    # �g���ƂȂ����t���[�Y����̂Ŗ��������
    CheckFileExists = $False;
    # ���݂��Ȃ��p�X���w�肳�ꂽ�Ƃ��x����\������
    CheckPathExists = $True;
    # STA�łȂ���OpenFileDialog���N�����Ȃ����A���L�̃I�v�V������L���ɂ���ƂȂ���MTA�ł������B�Ӗ��킩���
    ShowHelp = $True;
}

# �C�x���g
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
    # ExPing�̋N��
    $dt1 | ForEach {
        try
        {
            $ps = Start-Process -PassThru -FilePath  ($_.�p�X+"\ExPing.exe")
            $_.�v���Z�XID = $ps.ID
            #$global:plist += Start-Process -PassThru -FilePath  ($_.Path+"\ExPing.exe")
            $_.��� = "�N��"
        }
        catch [InvalidOperationException]
        {
            $_.��� = "�t�@�C����������܂���"
        }
    }
}
function PopupAutoRunning {
    #�x�����b�Z�[�W�̕\��
    $m = " �}�E�X����������āAEnter�������Ă��������B\n��������̊Ԃ̓}�E�X�𑀍삵�Ȃ��悤�ɂ��Ă�������"
    [System.Windows.Forms.MessageBox]::Show($m, "", "OK", "Information")
}
function RunPingExPing {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        # ExPing�̃R���g���[���擾
        #Write-Host ("�v���Z�XID�F"+$_.�v���Z�X)
        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F5);
        $_.��� = "Ping���s��"
    }
}
function StopPingExping {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        # ExPing�̃R���g���[���擾
        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick #| Out-Null #�Ȃ񂩂������x��
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F4); #����������
        $_.��� = "Ping��~"
    }
    #Start-Sleep -m 1000;

    # Ping���s���~�܂�܂őҋ@
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        While($true){
            $win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick | Out-Null
            $str = $win `
            | Get-UiaStatusBar -AutomationId 'StatusBar' -Class 'TStatusBar' `
            | Get-UiaEdit -AutomationId 'StatusBar.Pane0' `
            | Get-UiaEditText
            #Write-host $str
            if($str -eq '�I�����܂����B' ){ break }
            Start-Sleep -m 1000
        }
    }
    #Write-host "End Ping"
}
function SavePingExping {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        #�t�@�C���p�X�{�t�@�C����
        $save_dir1 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.�ۑ��t�@�C����+"_Ping����_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"
        $save_dir2 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.�ۑ��t�@�C����+"_Ping���v_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"
        #Write-Host $save_dir1

        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;

        # Ping���ʂ̕ۑ��_�C�A���O���J��
        <#
        $win | Get-UiaMenuBar -AutomationId 'MenuBar' -Name '�A�v���P�[�V����' `
             | Get-UiaMenuItem -AutomationId 'Item 1' -Name '�t�@�C��(F)' `
             | Invoke-UIAMenuItemExpand `
             | Get-UiaMenu -Class '#32768' -Name '�t�@�C��(F)' `
             | Get-UiaMenuItem -AutomationId 'Item 6' -Name 'Ping ���ʂ̕ۑ�(K)' `
             | Invoke-UIAMenuItemClick | Out-Null
        #>
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_K);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);

        Start-Sleep -m 1000;
        # �ۑ����邩�ǂ��������ꂽ���̃_�C�A���O���擾
        $sub = Get-UIAActivewindow
        # �t�@�C���������
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name '�t�@�C����(N):' `
        | Set-UIAEditText $save_dir1 | Out-Null
        # �ۑ��{�^��������
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '�ۑ�(S)' `
        | Invoke-UiaButtonClick | Out-Null

        # Ping���v�̕ۑ��_�C�A���O���J��
        <#
        $win | Get-UiaMenuBar -AutomationId 'MenuBar' -Name '�A�v���P�[�V����' `
             | Get-UiaMenuItem -AutomationId 'Item 1' -Name '�t�@�C��(F)' `
             | Invoke-UIAMenuItemExpand `
             | Get-UiaMenu -Class '#32768' -Name '�t�@�C��(F)' `
             | Get-UiaMenuItem -AutomationId 'Item 7' -Name 'Ping ���v�̕ۑ�(T)' `
             | Invoke-UIAMenuItemClick | Out-Null
        #>
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_T);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        Start-Sleep -m 1000;
        # �ۑ����邩�ǂ��������ꂽ���̃_�C�A���O���擾
        $sub = Get-UIAActivewindow
        # �t�@�C���������
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name '�t�@�C����(N):' `
        | Set-UIAEditText $save_dir2 | Out-Null
        # �ۑ��{�^��������
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '�ۑ�(S)' `
        | Invoke-UiaButtonClick | Out-Null
    }
}
function RunTraceExPing {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        # ExPing�̃R���g���[���擾
        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;
        #$win | Get-UiaToolBar -Class 'TToolBar' | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        #$win | Get-UiaButton -AutomationId 'Item 6' | Invoke-UiaButtonClick #| Out-Null
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::F6);
    }
    #Start-Sleep -m 1000;
}
function WaitTraceExPing {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        While($true){
            $win | Get-UiaButton -AutomationId 'Item 7' | Invoke-UiaButtonClick | Out-Null
            $str = $win `
            | Get-UiaStatusBar -AutomationId 'StatusBar' -Class 'TStatusBar' `
            | Get-UiaEdit -AutomationId 'StatusBar.Pane0' `
            | Get-UiaEditText
            #Write-host $str
            if($str -eq '�I�����܂����B' ){ break }
            Start-Sleep -m 1000
        }
    }
}
function SaveTraceExPing {
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {
        #�t�@�C���p�X�{�t�@�C����
        $save_dir3 = $cur_dir+"\"+$NumberTextBox.TEXT+"_"+$_.�ۑ��t�@�C����+"_Trace����_"+(Get-Date -Format "yyyyMMdd_HHmmss")+".csv"

        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;

        # Ping���ʂ̕ۑ��_�C�A���O���J��
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::SHIFT);
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::VK_K);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::SHIFT);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        Start-Sleep -m 1000;
        # �ۑ����邩�ǂ��������ꂽ���̃_�C�A���O���擾
        $sub = Get-UIAActivewindow
        # �t�@�C���������
        $sub `
        | Get-UiaEdit -AutomationId '1152' -Class 'Edit' -Name '�t�@�C����(N):' `
        | Set-UIAEditText $save_dir3 | Out-Null
        # �ۑ��{�^��������
        $sub `
        | Get-UiaButton -AutomationId '1' -Class 'Button' -Name '�ۑ�(S)' `
        | Invoke-UiaButtonClick | Out-Null
    }
}
function ClearResultExPing {
    #���ʂ��N���A����
    $dt1 | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} | ForEach {

        $win = Get-UiaWindow -ProcessId $_.�v���Z�XID;

        # Ping���ʂ̕ۑ��_�C�A���O���J��
        $win.Keyboard.KeyDown([WindowsInput.Native.VirtualKeyCode]::CONTROL);
        $win.Keyboard.KeyPress([WindowsInput.Native.VirtualKeyCode]::DELETE);
        $win.Keyboard.KeyUp([WindowsInput.Native.VirtualKeyCode]::CONTROL);
    }
}
function ShutdownExPing {
    $dt1 `
    | Where {$_.�v���Z�XID -ne [System.DBNULL]::Value} `
    | Where {(Get-Process -id $_.�v���Z�XID -ErrorAction SilentlyContinue) -ne $null} `
    | ForEach {
        Stop-Process -id $_.�v���Z�XID;
        $_.�v���Z�XID = [System.DBNULL]::Value
        $_.��� = "���N��"
    }
}
function read_csv {
    try
    {
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            Get-Content $ofd.FileName -ErrorAction Stop | ConvertFrom-Csv | ForEach {
                [void]$dt1.Rows.Add($_.No,$_.�p�X,$_.�ۑ��t�@�C����,$null,$null)
            }
        } else {
            exit
        }
    }
    catch [System.Management.Automation.ActionPreferenceStopException]
    {
        $str = $_.CategoryInfo.TargetName + "`n��������܂���"
        [void][System.Windows.Forms.MessageBox]::Show($str,"�G���[","OK","Warning" )
        read_csv
    }
}


read_csv
$form2.ShowDialog()
