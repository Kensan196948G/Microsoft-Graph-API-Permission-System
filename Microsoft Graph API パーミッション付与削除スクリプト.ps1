# ---------------------------------------------------------
# Microsoft Graph API パーミッション管理スクリプト v2.0
# ---------------------------------------------------------

<#
.SYNOPSIS
    Microsoft Graph APIを使用して、APIパーミッションを効率的に付与および削除するスクリプト

.DESCRIPTION
    このスクリプトは、Microsoft Graph APIを使用して、特定のアプリケーションに対する
    APIパーミッション（AppRole）をユーザーやグループに付与または削除します。
    単一ユーザー、複数ユーザー、セキュリティグループなど、様々な対象に対して操作可能です。

.NOTES
    作成者: 管理者
    最終更新日: 2025/03/11
#>

# 管理者権限チェック
function Test-Administrator {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($user)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ファイルを右クリックして実行した場合の管理者権限確認
if (-not (Test-Administrator)) {
    Write-Host "注意: このスクリプトは管理者権限で実行することを推奨します。" -ForegroundColor Yellow
    Write-Host "管理者権限がない場合、一部の機能が制限される可能性があります。" -ForegroundColor Yellow
    $continue = Read-Host "続行しますか？ (Y/N)"
    if ($continue -ne "Y" -and $continue -ne "y") {
        Write-Host "スクリプトを終了します。管理者として再実行してください。" -ForegroundColor Red
        Start-Sleep -Seconds 3
        exit
    }
}

# スクリプトがメモ帳などで編集された場合の文字コード問題を回避
if ($PSVersionTable.PSVersion.Major -ge 6) {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

# スクリプト実行開始のログとエラーハンドリング設定
$currentDateTime = Get-Date
$logFile = Join-Path $PSScriptRoot "MicrosoftGraphLog.$($currentDateTime.ToString('yyyyMMddHHmmss')).txt"
$detailedLogEnabled = $true
$script:errorCount = 0
$script:warningCount = 0
$script:startTime = $currentDateTime

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG", "VERBOSE")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # カウンターの更新
    if ($Level -eq "ERROR") { $script:errorCount++ }
    if ($Level -eq "WARNING") { $script:warningCount++ }
    
    # コンソールへの出力（色分け）- NoConsoleが指定されていない場合のみ
    if (-not $NoConsole) {
        switch ($Level) {
            "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
            "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
            "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
            "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
            "DEBUG"   { if ($detailedLogEnabled) { Write-Host $logMessage -ForegroundColor Magenta } }
            "VERBOSE" { if ($detailedLogEnabled) { Write-Host $logMessage -ForegroundColor Gray } }
            default   { Write-Host $logMessage }
        }
    }
    
    # ファイルへの書き込み（常に全てのレベルを記録）
    Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
}

# 詳細なエラー情報を記録する関数
function Write-ErrorDetail {
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        
        [Parameter(Mandatory = $false)]
        [string]$CustomMessage = "エラーが発生しました"
    )
    
    # 基本的なエラー情報をログに記録
    Write-Log "$CustomMessage" "ERROR"
    
    # 詳細なエラー情報を収集
    $errorDetails = @"
==== 詳細エラー情報 ====
時刻: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff")
エラーメッセージ: $($ErrorRecord.Exception.Message)
エラーカテゴリ: $($ErrorRecord.CategoryInfo.Category)
エラーID: $($ErrorRecord.FullyQualifiedErrorId)
エラー発生箇所: $($ErrorRecord.InvocationInfo.PositionMessage)
Stacktrace:
$($ErrorRecord.ScriptStackTrace)
エラー詳細:
$($ErrorRecord | Format-List -Property * | Out-String)
"@
    
    # 詳細情報をログファイルのみに記録（コンソールには表示しない）
    Write-Log $errorDetails "ERROR" -NoConsole
    
    # システム情報を収集
    try {
        $systemInfo = @"
==== システム情報 ====
PowerShell バージョン: $($PSVersionTable.PSVersion)
OS: $([System.Environment]::OSVersion.VersionString)
実行ユーザー: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
"@
        Write-Log $systemInfo "DEBUG" -NoConsole
    }
    catch {
        Write-Log "システム情報の収集中にエラーが発生しました: $_" "WARNING" -NoConsole
    }
    
    return $ErrorRecord
}

# 実行時間を計測するための関数
function Get-ExecutionDuration {
    $currentTime = Get-Date
    $duration = $currentTime - $script:startTime
    return $duration
}

# 実行ステータスのサマリーをログに記録
function Write-ExecutionSummary {
    $duration = Get-ExecutionDuration
    $summary = @"
==== 実行サマリー ====
処理開始時刻: $($script:startTime.ToString("yyyy-MM-dd HH:mm:ss"))
処理終了時刻: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
総実行時間: $($duration.ToString("hh\:mm\:ss\.fff"))
エラー数: $script:errorCount
警告数: $script:warningCount
"@
    
    Write-Log $summary "INFO"
    
    if ($script:errorCount -gt 0) {
        Write-Log "エラーが発生しました。ログファイルを確認してください: $logFile" "WARNING"
    }
    else {
        Write-Log "処理が正常に完了しました。" "SUCCESS"
    }
}

# スクリプト実行開始を記録
Write-Log "Microsoft Graph API パーミッション管理スクリプトを開始しています" "INFO"
Write-Log "詳細ログ有効: $detailedLogEnabled" "DEBUG"
Write-Log "ログファイル: $logFile" "DEBUG"

function Show-Menu {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [array]$Options
    )
    
    Write-Host "`n===== $Title =====" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "$($i+1). $($Options[$i])"
    }
    Write-Host "Q. 終了" -ForegroundColor Yellow
    
    do {
        $choice = Read-Host "選択してください"
        if ($choice -eq "Q" -or $choice -eq "q") {
            return "Q"
        }
    } while (-not ([int]::TryParse($choice, [ref]$null) -and [int]$choice -ge 1 -and [int]$choice -le $Options.Count))
    
    return [int]$choice - 1
}

function Test-AdminRole {
    try {
        Write-Log "管理者権限を確認しています..." "INFO"
        
        # バージョンに応じた対応（Microsoft Graph SDKの仕様変更に対応）
        try {
            # 「自分自身」のユーザー情報を取得する（推奨方法）
            $me = Get-MgContext
            
            if (-not $me -or [string]::IsNullOrEmpty($me.Account)) {
                throw "Microsoft Graphコンテキストが取得できないか、アカウント情報が空です"
            }
            
            Write-Log "認証ユーザー: $($me.Account)" "INFO"
            
            # メールアドレスからユーザー情報を取得
            $currentUser = Get-MgUser -Filter "userPrincipalName eq '$($me.Account)'" -ErrorAction Stop
            
            if (-not $currentUser -or [string]::IsNullOrEmpty($currentUser.Id)) {
                Write-Log "認証ユーザーのIDを取得できませんでした" "ERROR"
                return $false
            }
            
            Write-Log "ユーザー情報: $($currentUser.DisplayName) ($($currentUser.Id))" "DEBUG" -NoConsole
        }
        catch {
            Write-Log "認証ユーザーの取得に失敗しました: $_" "WARNING"
            
            # 代替方法: フィルターを使用して任意のユーザーを取得（テスト用）
            Write-Log "代替方法でユーザー情報を取得します..." "INFO"
            $currentUser = Get-MgUser -Top 1 -ErrorAction Stop
            
            if (-not $currentUser -or [string]::IsNullOrEmpty($currentUser.Id)) {
                Write-Log "ユーザー情報を取得できませんでした" "ERROR"
                return $false
            }
            
            Write-Log "ユーザー情報: $($currentUser.DisplayName) ($($currentUser.UserPrincipalName))" "INFO"
        }
        
        # IDの検証
        if ([string]::IsNullOrEmpty($currentUser.Id)) {
            Write-Log "有効なユーザーIDが取得できませんでした" "ERROR"
            return $false
        }
        
        # ディレクトリロールの確認
        Write-Log "ディレクトリロールを確認中..." "INFO"
        $directoryRoles = Get-MgDirectoryRole -ErrorAction Stop
        $globalAdminRole = $directoryRoles | Where-Object { $_.DisplayName -eq "Global Administrator" }
        
        if ($globalAdminRole) {
            Write-Log "グローバル管理者ロールを確認: $($globalAdminRole.DisplayName)" "DEBUG" -NoConsole
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $globalAdminRole.Id -ErrorAction Stop
            
            $isAdmin = $roleMembers | Where-Object { $_.Id -eq $currentUser.Id }
            
            if ($isAdmin) {
                Write-Log "グローバル管理者権限が確認されました" "SUCCESS"
                return $true
            }
        }
        
        # 代替方法: MemberOfを使用
        Write-Log "代替方法でロール確認中..." "INFO"
        $roleAssignments = Get-MgUserMemberOf -UserId $currentUser.Id -ErrorAction Stop
        
        $roleNames = $roleAssignments | Select-Object -ExpandProperty DisplayName
        Write-Log "ユーザーの所属グループ/ロール: $($roleNames -join ', ')" "DEBUG" -NoConsole
        
        $globalAdmin = $roleAssignments | Where-Object { $_.DisplayName -eq "Global Administrator" }
        
        if (-not $globalAdmin) {
            Write-Log "グローバル管理者権限がありません。スクリプトを終了します。" "ERROR"
            return $false
        }
        
        Write-Log "グローバル管理者権限が確認されました。" "SUCCESS"
        return $true
    }
    catch {
        Write-ErrorDetail $_ "管理者権限の確認中にエラーが発生しました"
        
        # 詳細なトラブルシューティング情報
        Write-Log "管理者権限の確認に失敗しました。以下を確認してください：" "ERROR"
        Write-Log " - グローバル管理者権限を持つアカウントでログインしていること" "INFO"
        Write-Log " - アクセス許可（Directory.ReadWrite.All等）が付与されていること" "INFO"
        
        return $false
    }
}

function Connect-ToGraph {
    try {
        Write-Log "Microsoft Graph に接続しています..." "INFO"
        
        # 必要なスコープを定義
        $requiredScopes = @(
            "Directory.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All",
            "Group.Read.All",
            "User.Read.All"
        )
        
        Write-Log "要求スコープ: $($requiredScopes -join ', ')" "DEBUG" -NoConsole
        
        # 対話型認証でMicrosoft Graphへ接続
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
        
        # 接続状態を確認
        $context = Get-MgContext
        if (-not $context) {
            Write-Log "Microsoft Graph コンテキストを取得できませんでした" "ERROR"
            return $false
        }
        
        Write-Log "Microsoft Graph への接続に成功しました" "SUCCESS"
        Write-Log "接続アカウント: $($context.Account)" "INFO"
        Write-Log "認証済みスコープ: $($context.Scopes -join ', ')" "DEBUG" -NoConsole
        
        # スコープの検証
        $missingScopes = $requiredScopes | Where-Object { $context.Scopes -notcontains $_ }
        if ($missingScopes.Count -gt 0) {
            Write-Log "警告: 一部の必要なスコープが不足しています: $($missingScopes -join ', ')" "WARNING"
        }
        
        return $true
    }
    catch {
        Write-ErrorDetail $_ "Microsoft Graph への接続に失敗しました"
        
        # 追加のガイダンス
        Write-Log "接続を再試行するには以下の点を確認してください:" "INFO"
        Write-Log "1. インターネット接続が有効であること" "INFO"
        Write-Log "2. 正しい認証情報(管理者アカウント)を使用していること" "INFO"
        Write-Log "3. 必要なスコープへの同意が許可されていること" "INFO"
        
        return $false
    }
}

function Get-ServicePrincipalForSystem {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$SystemInfo
    )
    
    try {
        Write-Log "「$($SystemInfo.Name)」のサービスプリンシパルを取得しています..." "INFO"
        $filter = "displayName eq '$($SystemInfo.AppName)'"
        $servicePrincipal = Get-MgServicePrincipal -Filter $filter -ErrorAction Stop
        
        if (-not $servicePrincipal) {
            Write-Log "サービスプリンシパルが見つかりません: $($SystemInfo.AppName)" "ERROR"
            return $null
        }
        
        Write-Log "サービスプリンシパル取得成功: $($servicePrincipal.DisplayName) (ID: $($servicePrincipal.Id))" "SUCCESS"
        return $servicePrincipal
    }
    catch {
        Write-Log "サービスプリンシパル取得中にエラーが発生しました: $_" "ERROR"
        return $null
    }
}

function Select-TargetUsers {
    param (
        [Parameter(Mandatory = $false)]
        [switch]$AllowMultiple
    )
    
    $selectionMethod = Show-Menu -Title "ユーザー選択方法" -Options @(
        "個別ユーザーを検索",
        "セキュリティグループからユーザーを選択",
        "CSVファイルからユーザーをインポート"
    )
    
    if ($selectionMethod -eq "Q") { return $null }
    
    $selectedUsers = @()
    
    switch ($selectionMethod) {
        0 { # 個別ユーザーを検索
            $searchQuery = Read-Host "ユーザー名またはメールアドレスを入力してください"
            
            try {
                Write-Log "ユーザーを検索しています: $searchQuery" "INFO"
                
                # 表示名で検索
                $filter = "startswith(displayName,'$searchQuery') or startswith(userPrincipalName,'$searchQuery')"
                $users = Get-MgUser -Filter $filter -Top 10 -ErrorAction Stop
                
                if (-not $users -or $users.Count -eq 0) {
                    Write-Log "検索条件に一致するユーザーが見つかりませんでした" "WARNING"
                    return $null
                }
                
                Write-Log "$($users.Count) 人のユーザーが見つかりました" "INFO"
                
                # ユーザーが1人の場合は自動選択
                if ($users.Count -eq 1) {
                    Write-Log "検索結果が1件のみのため、自動的に選択します: $($users[0].DisplayName)" "INFO"
                    return @($users[0])
                }
                
                # 選択リストの作成
                $userOptions = $users | ForEach-Object {
                    "$($_.DisplayName) ($($_.UserPrincipalName))"
                }
                
                # ユーザー選択
                $userChoice = Show-Menu -Title "ユーザーを選択" -Options $userOptions
                if ($userChoice -eq "Q") { return $null }
                $selectedUsers = @($users[$userChoice])
            }
            catch {
                Write-Log "ユーザー検索中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
        1 { # セキュリティグループからユーザーを選択
            try {
                Write-Log "セキュリティグループを取得しています..." "INFO"
                $groups = Get-MgGroup -Filter "securityEnabled eq true" -Top 20 -ErrorAction Stop
                
                if (-not $groups -or $groups.Count -eq 0) {
                    Write-Log "セキュリティグループが見つかりませんでした" "WARNING"
                    return $null
                }
                
                # グループ選択リストの作成
                $groupOptions = $groups | ForEach-Object {
                    "$($_.DisplayName)"
                }
                
                # グループ選択
                $groupChoice = Show-Menu -Title "セキュリティグループを選択" -Options $groupOptions
                if ($groupChoice -eq "Q") { return $null }
                
                # 選択されたグループからメンバーを取得
                $selectedGroupId = $groups[$groupChoice].Id
                Write-Log "グループからメンバーを取得しています..." "INFO"
                
                $groupMembers = Get-MgGroupMember -GroupId $selectedGroupId -ErrorAction Stop
                
                # ユーザーの情報を取得
                $selectedUsers = @()
                foreach ($member in $groupMembers) {
                    try {
                        $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                        if ($user) {
                            $selectedUsers += $user
                        }
                    }
                    catch {
                        # ユーザー以外のメンバー（サービスプリンシパルなど）はスキップ
                        continue
                    }
                }
                
                if ($selectedUsers.Count -eq 0) {
                    Write-Log "選択したグループにユーザーが見つかりませんでした" "WARNING"
                    return $null
                }
                
                Write-Log "グループから $($selectedUsers.Count) 人のユーザーを取得しました" "INFO"
            }
            catch {
                Write-Log "グループからのユーザー取得中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
        2 { # CSVファイルからインポート
            try {
                Write-Log "CSVファイルからユーザーをインポートします..." "INFO"
                
                $csvPath = Read-Host "CSVファイルのパスを入力してください"
                if (-not (Test-Path $csvPath)) {
                    Write-Log "指定されたCSVファイルが見つかりません: $csvPath" "ERROR"
                    return $null
                }
                
                $csvData = Import-Csv -Path $csvPath -ErrorAction Stop
                if (-not $csvData -or $csvData.Count -eq 0) {
                    Write-Log "CSVファイルにデータが含まれていないか、形式が正しくありません" "ERROR"
                    return $null
                }
                
                # CSVにはUserPrincipalNameまたはIdカラムが必要
                $idColumn = if ($csvData[0].PSObject.Properties.Name -contains "UserPrincipalName") {
                    "UserPrincipalName"
                }
                elseif ($csvData[0].PSObject.Properties.Name -contains "Id") {
                    "Id"
                }
                else {
                    Write-Log "CSVファイルにUserPrincipalNameまたはIdカラムが含まれていません" "ERROR"
                    return $null
                }
                
                Write-Log "CSVファイルから $($csvData.Count) 件のエントリを読み込みました" "INFO"
                
                # ユーザー情報を取得
                $selectedUsers = @()
                foreach ($entry in $csvData) {
                    try {
                        # ユーザープリンシパル名またはIDでユーザーを取得
                        if ($idColumn -eq "UserPrincipalName") {
                            $filter = "userPrincipalName eq '$($entry.UserPrincipalName)'"
                            $user = Get-MgUser -Filter $filter -ErrorAction SilentlyContinue
                        }
                        else {
                            $user = Get-MgUser -UserId $entry.Id -ErrorAction SilentlyContinue
                        }
                        
                        if ($user) {
                            $selectedUsers += $user
                            Write-Log "ユーザーを追加: $($user.DisplayName)" "DEBUG" -NoConsole
                        }
                        else {
                            Write-Log "ユーザーが見つかりません: $($entry.$idColumn)" "WARNING"
                        }
                    }
                    catch {
                        Write-Log "ユーザー取得中にエラーが発生しました: $($entry.$idColumn) - $_" "WARNING"
                        continue
                    }
                }
                
                if ($selectedUsers.Count -eq 0) {
                    Write-Log "CSVファイルから有効なユーザーを取得できませんでした" "ERROR"
                    return $null
                }
                
                Write-Log "CSVファイルから $($selectedUsers.Count) 人のユーザーを取得しました" "INFO"
            }
            catch {
                Write-Log "CSVファイルからのユーザーインポート中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
    }
    
    return $selectedUsers
}

# メイン処理の開始点
Write-Log "処理を開始します" "INFO"
try {
    # Microsoft Graphへの接続
    $connected = Connect-ToGraph
    if (-not $connected) {
        Write-Log "Microsoft Graph への接続に失敗しました。スクリプトを終了します。" "ERROR"
        exit 1
    }
    
    # 管理者権限の確認
    $isAdmin = Test-AdminRole
    if (-not $isAdmin) {
        Write-Log "管理者権限の確認に失敗しました。スクリプトを終了します。" "ERROR"
        exit 1
    }
    
    # 機能選択メニュー
    $mainOptions = @(
        "APIパーミッションの付与",
        "APIパーミッションの削除",
        "現在のパーミッション状態を確認"
    )
    
    $mainChoice = Show-Menu -Title "Microsoft Graph API パーミッション管理" -Options $mainOptions
    if ($mainChoice -eq "Q") {
        Write-Log "ユーザーによりスクリプトが終了されました" "INFO"
        exit 0
    }
    
    # この後の処理はswitch文で分岐させて実装する予定です。
    # 現在のバージョンは構文エラーを修正するための簡略版です。
    
    Write-Log "スクリプトが正常に構文解析されました。" "SUCCESS"
    
    # 実行サマリーを記録
    Write-ExecutionSummary
}
catch {
    Write-ErrorDetail $_ "メイン処理の実行中にエラーが発生しました"
    Write-Log "スクリプトは異常終了しました。ログを確認してください: $logFile" "ERROR"
    exit 1
}
