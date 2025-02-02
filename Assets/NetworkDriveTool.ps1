#Requires -Version 5.0
<#
 .Synopsis
   ネットワークドライブのマウント

 .Description
   リモートフォルダーをネットワークドライブとしてマウントします
   サーバーへの接続ができない (ping応答が4回中2回失敗) 場合、30秒後に再接続を試みます

 .Notes
   2025-02-02 更新

 .Parameter DriveLetter
   ドライブ文字

 .Parameter Root
   リモートフォルダーのUNCパス
   
 .Parameter TestServerName
   Pingテストをするサーバー名やIPアドレス

 .Parameter DriveLabel
   ドライブ名

 .Parameter BeforeDismountDrive
   事前にドライブの接続を解除

 .Parameter IsLogging
   %Temp%\MountNetworkTool.log にログを追記
   ログの削除は本スクリプトでは実施しないため、常用はしないこと

 .Example
   最小
   PS> MountNetworkTool.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv"

 .Example
   ドライブ名指定あり
   PS> MountNetworkTool.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv" -DriveLabel "共有"

 .Example
   認証情報をログイン後に指定する
   PS> MountNetworkTool.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv" -DriveLabel "共有" -BeforeDismountDrive

 .Example
   コンソールウィンドウホスト を明示的に指定し、Windows ターミナルからの実行を迂回・ウィンドウを非表示にする
   C:\Windows\System32\conhost.exe PowerShell -ExecutionPolicy ByPass -WindowStyle Hidden -File MountNetworkTool.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv" -DriveLabel "共有"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$True)] [Char] $DriveLetter,
    [Parameter(Mandatory=$True)] [String] $Root,
    [Parameter(Mandatory=$True)] [String] $TestServerName,
    [Parameter()] [String] $DriveLabel,
    [Parameter()] [Switch] $BeforeDismountDrive=$False,
    [Parameter()] [Switch] $IsLogging=$False
)

If ($IsLogging){
    Start-Transcript -Path (Join-Path $env:TEMP "MountNetworkTool.log") -Append
}

Function Test-ServerConnection(){
    # 4回中2回成功したら待機無し
    Write-Host "$(Get-Date -Format F): サーバーへ接続しています"
    $PingResult = Test-Connection -ComputerName $TestServerName -Count 4 -Delay 1 -ErrorAction Ignore | Where-Object { $_.StatusCode -eq 0 }
    If ($PingResult.Count -lt 2) {
        Write-Warning "$(Get-Date -Format F): サーバーからの応答が無いため30秒待機します"
        Start-Sleep -Seconds 30
        Test-ServerConnection
    }
    Write-Host "$(Get-Date -Format F): サーバーへの接続に成功しました"
    Mount-NetworkDrive
}

Function Test-NetworkDrive(){
    $DriveInfo = Get-PSDrive | Where-Object Name -eq $DriveLetter
    If ($DriveInfo.Count -eq 0){
        Return $False
    }
    Else{
        Return ($DriveInfo.DisplayRoot.ToLower() -eq $Root.ToLower() -and $DriveInfo.Name.ToLower() -eq $DriveLetter.Tostring().ToLower())
    }
}

Function Dismount-NetworkDrive(){
    Try{
        If ((Get-PSDrive | Where-Object Name -eq $DriveLetter).Count){
            Write-Host "$(Get-Date -Format F): ドライブのマウントを解除しています"
            #Remove-SmbMapping "$($DriveLetter):" -Force | Out-Null
            Start-Process net.exe "use $($DriveLetter): /delete" -Wait -WindowStyle Hidden
            $Count = 0
            while($Count -ne 15)
            {
                $DriveInfo = Get-PSDrive | Where-Object Name -eq $DriveLetter
                If ($DriveInfo.Count -eq 0){
                    Break
                }
                $Count++
                Start-Sleep -Seconds 1
                Write-Host "$(Get-Date -Format F): ドライブのマウント解除を再試行しています"
            }
        }
    }
    Catch{
        Write-Warning "$(Get-Date -Format F): ドライブのマウント解除に失敗しました`n$($_.Exception.Message)"
    }
}

Function Mount-NetworkDrive(){
    If ($BeforeDismountDrive){
        Dismount-NetworkDrive
    }
    If (Test-NetworkDrive){
        Write-Host "$(Get-Date -Format F): すでにマウントされています"
    }
    Else{
        Dismount-NetworkDrive
        Try{
            New-PSDrive -Name $DriveLetter -Root $Root -PSProvider FileSystem -Persist -Scope Global -ErrorAction Stop | Out-Null
            #New-SmbMapping が動作しない環境があった
            #New-SmbMapping -LocalPath $DriveLetter -RemotePath $Root -Persistent $True -ErrorAction Stop | Out-Null
        }
        Catch{
            Write-Warning "$(Get-Date -Format F): ドライブのマウントができませんでしたので再試行します`n$($_.Exception.Message)"
            Start-Sleep -Seconds 5
            Test-ServerConnection
        }
    }
    If (-not [String]::IsNullOrWhiteSpace($DriveLabel)){
        Try{
            Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\$($Root.Replace("\","#"))" -Name "_LabelFromReg" -Value $DriveLabel -Force -ErrorAction Stop
        }
        Catch{
            Write-Warning "$(Get-Date -Format F): ドライブ名の指定に失敗`n$($_.Exception.Message)"
        }
    }
    Start-Sleep -Seconds 5
    If (Test-NetworkDrive){
        Write-Host "$(Get-Date -Format F): ドライブのマウントに成功しました"
    }
    Else{
        Write-Warning "$(Get-Date -Format F): ドライブのマウントに失敗した可能性があります"
    }
}

Write-Host "Root: $Root`nDriveLetter: $DriveLetter`nTestServerName: $TestServerName`nDriveLabel: $DriveLabel`nBeforeDismountDrive`nBeforeDismountDrive: $BeforeDismountDrive"
Test-ServerConnection

If ($IsLogging){
    Stop-Transcript
}
