#Requires -Version 5.0
<#
 .Synopsis
   ネットワークドライブのマウント

 .Description
   リモートフォルダーをネットワークドライブとしてマウントします

 .Notes
   2025-01-26 更新

 .Parameter DriveLetter
   ドライブ文字

 .Parameter Root
   リモートフォルダーのUNCパス
   
 .Parameter TestServerName
   Pingテストをするサーバー名やIPアドレス

 .Parameter DriveLabel
   ドライブ名

 .Parameter IsLogging
   %Temp%\NetworkDriveMount.log にログを追記します

 .Example
   # 最小
   MountNetworkDrive.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv"
   # ドライブ名指定あり
   MountNetworkDrive.ps1 -DriveLetter "Z" -Root "\\file-sv\share" -TestServerName "file-sv" -DriveLabel "共有"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$True)] [Char] $DriveLetter,
    [Parameter(Mandatory=$True)] [String] $Root,
    [Parameter(Mandatory=$True)] [String] $TestServerName,
    [Parameter()] [String] $DriveLabel,
    [Parameter()] [String] $IsLogging=$False
)

If ($IsLogging){
    Start-Transcript -Path (Join-Path $env:TEMP "NetworkDriveMount.log") -Append
}

Function Test-ServerConnection(){
    # 4回中2回成功したら待機無し
    Write-Host "$(Get-Date -Format F): サーバーへ接続しています"
    $PingResult = Test-Connection -ComputerName $TestServerName -Count 4 -Delay 1 -ErrorAction Ignore | Where-Object { $_.StatusCode -eq 0 }
    If ($PingResult.Count -lt 2) {
        Write-Warning "$(Get-Date -Format F): [サーバーからの応答が無いため30秒待機します] $TestServerName"
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

Function Mount-NetworkDrive(){
    If (Test-NetworkDrive){
        Write-Host "$(Get-Date -Format F): すでにマウントされています"
    }
    Else{
        Try{
            If ((Get-PSDrive | Where-Object Name -eq $DriveLetter).Count){
                Write-Host "$(Get-Date -Format F): ドライブのマウントを解除しています"
                Remove-SmbMapping "$($DriveLetter):"
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

Write-Host "Root: $Root`nDriveLetter: $DriveLetter`nTestServerName: $TestServerName`nDriveLabel: $DriveLabel`n"
Test-ServerConnection
If ($IsLogging){
    Stop-Transcript
}
