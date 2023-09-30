

# // カレントディレクトリの取得
$currentDir = $PSScriptRoot
# // ログファイル
$destLogPath = Join-Path $currentDir "Log"
# // 端末リストファイル
$pcListFileName = "PCLists.txt"
# // 情報を取得したいアプリケーションのリストファイル
$appListFileName = "TargetApp.txt"
# // ファイル/フォルダの確認
if(-not (Test-Path $destLogPath)) {Write-Output "Logフォルダが存在しません。";exit}
if(-not (Test-Path $pcListFileName)) {Write-Output "$($pcListFileName)ファイルが存在しません。";exit}
if(-not (Test-Path $appListFileName)) {Write-Output "$($appListFileName)ファイルが存在しません。";exit}
# // 端末リストを読み込み
$pcList = @(Get-Content $pcListFileName)
# // 取得するAppリストを読み込み
$appList = @(Get-Content $appListFileName)
# // スクリプトファイル名(拡張子無し)
$scriptNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
# // 資材実行の日付
$today = (Get-Date -Format "yyyyMMdd")

# // ログ出力
function Output-LogFile 
{
    param([String]$logPath, [PSCustomObject]$outputMessage)
    $outputMessage | ConvertTo-Csv -NoTypeInformation | Set-Content $logPath
}

# // 対象アプリのバージョンをPSCustomObjectへ格納
function Set-AppVersion
{
    param([PSCustomObject]$appObject,$instApp="-")
    if($instApp -is [array])
    {
        foreach($app in $appList)
        {
            $version = ($instApp | ?{$_.Name -like $app}).Version
            $appObject | Add-Member -MemberType NoteProperty -Name $app.ToString() -Value $version
        }
    }
    else
    {
        foreach($app in $appList)
        {
            $appObject | Add-Member -MemberType NoteProperty -Name $app.ToString() -Value $instApp
        }
    }
    return $appObject
}

# // ログ出力用のPSCustomObjectの作成
function Make-PSCustomObject
{
    param([String]$pc,[String]$message="")
    $tempList = @()
    $tempList = [PSCustomObject]@{
        day = $today
        pc = $pc
        message = $message
    }
    if("" -eq $message)
    {
        # // インストールされているアプリケーションを取得
        $instLists = (Get-WmiObject -Class Win32_Product -ComputerName $pc)
        $tempList = Set-AppVersion $tempList $instLists
    }
    else
    {
        # // アプリケーション情報の取得に失敗/ping不可のケース
        $tempList  = Set-AppVersion $tempList
    }
    return $tempList
}

# // 読み込んだ端末リストから1台チェック
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
        $logPath_pcName = Join-Path $destLogPath ($today,$pcName,($scriptNameWithoutExtension + ".csv") -join "_")
        # // 疎通確認
        $resultNetwork = Test-Connection $pcName -Count 1 -Quiet
        if(-not $resultNetwork)
        {
            # // pingがNGの処理
            $list = Make-PSCustomObject $pcName "pingNG"
            Output-LogFile $logPath_pcName $list
            Continue
        }
        # // アプリケーションのバージョン情報を取得
        $list = Make-PSCustomObject $pcName
        Output-LogFile $logPath_pcName $list
    }
    catch
    {
        # // エラー発生時の処理
        $list = Make-PSCustomObject $pcName "エラーが発生しました"
        Output-LogFile $logPath_pcName $list
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
Sammary-Log ("Sammary_" + $scriptNameWithoutExtension + ".csv") $destLogPath

exit

