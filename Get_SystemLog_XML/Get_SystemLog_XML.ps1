

# // カレントディレクトリの取得
$currentDir = $PSScriptRoot
# // ログファイル
$destLogPath = Join-Path $currentDir "Log"
# // 端末リストファイル
$pcListFileName = "PCLists.txt"
# // CSVファイルでログを出力
$logFileName = "SystemLog_Logon.csv"
# // 抽出条件のXMLファイル
# //   イベントビューアーの「現在のログをフィルター」から条件を記入して、「XML」タブからQueryを取得し、xmlファイルにコピペ
$xmlFile = "QueryList.xml"
# // ファイル/フォルダの確認
if(-not (Test-Path $destLogPath)) {Write-Output "Logフォルダが存在しません。";exit}
if(-not (Test-Path $pcListFileName)) {Write-Output "$($pcListFileName)ファイルが存在しません。";exit}
if(-not (Test-Path $xmlFile)) {Write-Output "$($xmlFile)ファイルが存在しません。";exit}
# // 端末リストを読み込み
$pcList = @(Get-Content $pcListFileName)
# // xmlファイルを読み込み、xml化して変数に格納
$xmlQuery = [XML]@(Get-Content $xmlFile)

# // ログ出力
function Output-LogFile 
{
    param([String]$logPath, [PSCustomObject]$outputMessage)
    $outputMessage | ConvertTo-Csv -NoTypeInformation | Set-Content $logPath
}
# // リモートレジストリサービスの起動
function Set_RemoteRegistry 
{
    param([String]$pc)
    # // リモートレジストリを自動起動
    Set-Service -Name RemoteRegistry -StartupType Automatic -ComputerName $pc | Out-Null
}
# // リモートレジストリサービスの起動確認
function Check_RemoteRegistry 
{
    param([String]$pc)
    $serviceRemoteRegistry = Get-WmiObject Win32_Service -Filter "Name='RemoteRegistry'" -ComputerName $pc
    if($serviceRemoteRegistry.StartMode -ne "Auto")
    {
        # // リモートレジストリを自動起動
        Ser_RemoteRegistry $pc
    }
}
# // システムログにてログオンユーザーの情報を取得
function Get_UserNameFromSID
{
    param([String]$pc,[String]$sid)
    # // リモートレジストリサービスの起動確認
    Check_RemoteRegistry $pc
    $regHive = "LocalMachine"
    $regView = "Registry64"
    # // レジストリパス
    $regPath = Join-Path "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" $sid
    $key = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::$regHive,$pc,[Microsoft.Win32.RegistryView]::$regView)
    $regKey = $key.OpenSubKey($regPath)
    $tempUserName = $regKey.GetValue("ProfileImagePath")
    # // ユーザープロファイルからユーザー部分のみ取得 (C:\Users\Tomo　⇒　Tomo)
    $tempUserName = $tempUserName.ToString() -split "\\"
    $userName = $tempUserName[-1]
    return $userName
}
# // ログ出力用のPSCustomObjectの作成
function Make-PSCustomObject
{
    param([String]$pc,[String]$message="",[Array]$events="")
    $tempList = @()
    if("" -eq $message)
    {
        # // イベント情報をPSCustomObjectへ格納
        foreach($event in $events)
        {
            # // イベントをXML化
            $eventXML = [XML]($event.ToXml())
            $tempList += [PSCustomObject]@{
                pcName = $pc
                message = $event.Message
                Id = $event.Id
                Time = $event.TimeCreated
                #SID = $event.properties[1].Value
                SID = ($eventXML.Event.EventData.Data | ?{$_.Name -eq "UserSid"})."#Text"
                UserName = Get_UserNameFromSID -pc $pc -sid ($eventXML.Event.EventData.Data | ?{$_.Name -eq "UserSid"})."#Text"
            }
        }
    }
    else
    {
        # // イベント情報の取得に失敗/ping不可のケース
        $tempList = [PSCustomObject]@{
            pcName = $pc
            message = $message
            Id = "-"
            Time = "-"
            SID = "-"
            UserName = "-"
        }
    }
    return $tempList
}

# // ログ取得
foreach ($pcName in $pcList)
{
    try
    {
        # // エラーがあった場合処理を停止
        $ErrorActionPreference = "Stop"
        # // 端末リストに端末名が無い場合(空行)は、スキップ
        if("" -eq $pcName) { Continue }
        $list = @()
        # // 端末ごとのログ出力先
        $logPath_pcName = Join-Path $destLogPath ($pcName,$logFileName -join "_")

        # // 疎通確認
        if(-Not (Test-Connection $pcName -Count 1 -Quiet))
        {
            # // pingがNGの処理
            $list = Make-PSCustomObject -pc $pcName -message "pingNG"
            Output-LogFile -logPath $logPath_pcName -outputMessage $list
            Continue
        }
        # // イベントログの取得
        #       条件のログが存在しない場合のために、SilentlyContinue を指定
        $systemEvents = Get-WinEvent -ComputerName $pcName -FilterXml $xmlQuery -MaxEvents 10 -ErrorAction SilentlyContinue
        # // 取得できるイベントログが存在しない場合は、処理を終了
        if($null -eq $systemEvents)
        {
            $list = Make-PSCustomObject -pc $pcName -message "条件のログが存在しませんでした"
            Output-LogFile -logPath $logPath_pcName -outputMessage $list
        }
        else
        {
            # // 取得したイベントログを処理
            $list = Make-PSCustomObject -pc $pcName -events $systemEvents
            Output-LogFile -logPath $logPath_pcName -outputMessage $list
        }
    }
    catch
    {
        # // エラー発生時の処理
        $list = Make-PSCustomObject -pc $pcName -message "エラーが発生しました"
        Output-LogFile -logPath $logPath_pcName -outputMessage $list
    }
}


# // ログをファイルに出力した際に使用
function Sammary-Log
{
    param([String]$fileName,[String]$targetLogPath)
    # // ログ格納先のファイル情報を取得
    $files = @(Get-ChildItem $targetLogPath *.csv).FullName
    # // サマリーログを作成
    $sammaryLogPath = Join-Path $currentDir $fileName
    Import-Csv $files -Encoding Default | Export-Csv $sammaryLogPath -Encoding Default -NoTypeInformation
}
Sammary-Log "Sammary_SystemLog_Logon.csv" $destLogPath

exit

