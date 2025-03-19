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
            # ユーザー検索の再試行ループ
            $searchSuccess = $false
            while (-not $searchSuccess) {
                $searchQuery = Read-Host "ユーザー名、メールアドレス、またはSAMアカウント名を入力してください"
                
                # 入力確認
                $confirmed = $false
                while (-not $confirmed) {
                    $confirm = Read-Host "入力内容「$searchQuery」で検索しますか？ (Y/N)"
                    if ($confirm -eq "Y" -or $confirm -eq "y") {
                        $confirmed = $true
                    }
                    elseif ($confirm -eq "N" -or $confirm -eq "n") {
                        $searchQuery = Read-Host "ユーザー名、メールアドレス、またはSAMアカウント名を入力してください"
                    }
                    else {
                        Write-Host "Y または N を入力してください" -ForegroundColor Yellow
                    }
                }
                
                try {
                    Write-Log "ユーザーを検索しています: $searchQuery" "INFO"
                    
                    # 複数の検索条件を組み合わせる
                    # 1. 表示名での検索
                    # 2. メールアドレスでの検索
                    # 3. SAMアカウント名での検索
                    $filter = "startswith(displayName,'$searchQuery') or startswith(userPrincipalName,'$searchQuery')"
                    $users = Get-MgUser -Filter $filter -Top 10 -ErrorAction Stop
                    
                    # SAMアカウント名での検索を追加
                    # OnPremisesSamAccountNameは直接フィルタできないため、別途取得
                    if (-not $users -or $users.Count -eq 0) {
                        Write-Log "表示名・メールアドレスで見つからなかったため、SAMアカウント名で検索します" "INFO"
                        
                        # 最初に1000人程度のユーザーを取得
                        $allUsers = Get-MgUser -Top 1000 -Property DisplayName, UserPrincipalName, OnPremisesSamAccountName, Id -ErrorAction Stop
                        
                        # SAMアカウント名でフィルタリング
                        $users = $allUsers | Where-Object { 
                            $_.AdditionalProperties.onPremisesSamAccountName -and 
                            $_.AdditionalProperties.onPremisesSamAccountName.StartsWith($searchQuery)
                        }
                    }
                    
                    if (-not $users -or $users.Count -eq 0) {
                        Write-Log "検索条件に一致するユーザーが見つかりませんでした" "WARNING"
                        $retry = Read-Host "再検索しますか？ (Y/N)"
                        if ($retry -ne "Y" -and $retry -ne "y") {
                            return $null
                        }
                        # ループ継続（再検索）
                        continue
                    }
                    
                    Write-Log "$($users.Count) 人のユーザーが見つかりました" "INFO"
                    
                    # ユーザーが1人の場合は確認後に選択
                    if ($users.Count -eq 1) {
                        $samAccountName = "N/A"
                        if ($users[0].AdditionalProperties -and $users[0].AdditionalProperties.onPremisesSamAccountName) {
                            $samAccountName = $users[0].AdditionalProperties.onPremisesSamAccountName
                        }
                        
                        $userDisplayName = $users[0].DisplayName
                        if ([string]::IsNullOrEmpty($userDisplayName)) {
                            $userDisplayName = "(表示名なし)"
                        }
                        
                        $userUPN = $users[0].UserPrincipalName
                        
                        Write-Host "見つかったユーザー: $userDisplayName ($userUPN) [SAM: $samAccountName]" -ForegroundColor Cyan
                        $confirm = Read-Host "このユーザーを選択しますか？ (Y/N)"
                        
                        if ($confirm -eq "Y" -or $confirm -eq "y") {
                            Write-Log "ユーザーを選択しました: $userDisplayName ($userUPN)" "INFO"
                            $searchSuccess = $true
                            return @($users[0])
                        }
                        else {
                            # ループ継続（再検索）
                            continue
                        }
                    }
                    
                    # 複数ユーザーの場合は選択肢を表示
                    # 選択リストにSAMアカウント名を含める
                    $usersWithSAM = @()
                    foreach ($user in $users) {
                        $samAccountName = "N/A"
                        if ($user.AdditionalProperties.onPremisesSamAccountName) {
                            $samAccountName = $user.AdditionalProperties.onPremisesSamAccountName
                        }
                        
                        $usersWithSAM += [PSCustomObject]@{
                            DisplayName = $user.DisplayName
                            UserPrincipalName = $user.UserPrincipalName
                            SamAccountName = $samAccountName
                            User = $user
                        }
                    }
                    
                    # 表形式で表示
                    Write-Host "`n===== 検索結果 =====" -ForegroundColor Cyan
                    $format = "{0,-3} | {1,-20} | {2,-30} | {3,-15}"
                    Write-Host ($format -f "No.", "表示名", "UPN", "SAMアカウント名")
                    Write-Host ("-" * 75)
                    
                    for ($i = 0; $i -lt $usersWithSAM.Count; $i++) {
                        # 表示名を処理（長すぎる場合は省略）
                        $displayName = $usersWithSAM[$i].DisplayName
                        if ($displayName.Length -gt 18) {
                            $displayName = $displayName.Substring(0, 15) + "..."
                        }
                        
                        # UPNを処理（長すぎる場合は省略）
                        $upn = $usersWithSAM[$i].UserPrincipalName
                        if ($upn.Length -gt 28) {
                            $upn = $upn.Substring(0, 25) + "..."
                        }
                        
                        # 表示
                        Write-Host ($format -f ($i+1), $displayName, $upn, $usersWithSAM[$i].SamAccountName)
                    }
                    
                    # ユーザー選択
                    $userOptions = $users | ForEach-Object {
                        $sam = if ($_.AdditionalProperties.onPremisesSamAccountName) { $_.AdditionalProperties.onPremisesSamAccountName } else { "N/A" }
                        "$($_.DisplayName) ($($_.UserPrincipalName)) [SAM: $sam]"
                    }
                    
                    $userChoice = Show-Menu -Title "ユーザーを選択" -Options $userOptions
                    if ($userChoice -eq "Q") { return $null }
                    
                    $selectedUsers = @($users[$userChoice])
                    Write-Log "ユーザーを選択しました: $($selectedUsers[0].DisplayName)" "INFO"
                    $searchSuccess = $true
                }
                catch {
                    Write-Log "ユーザー検索中にエラーが発生しました: $_" "ERROR"
                    $retry = Read-Host "再検索しますか？ (Y/N)"
                    if ($retry -ne "Y" -and $retry -ne "y") {
                        return $null
                    }
                    # ループ継続（再検索）
                }
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

# システム情報の定義
$systemDefinitions = @{
    "OneDrive" = @{
        Name = "OneDrive for Business"
        AppName = "Office 365 SharePoint Online"
        AppRoles = @(
            @{
                Name = "User.Read.All"
                Description = "ユーザー情報へのアクセス"
            },
            @{
                Name = "Directory.Read.All"
                Description = "組織構造情報へのアクセス"
            },
            @{
                Name = "Files.ReadWrite.All"
                Description = "OneDriveファイルの読み書き"
            }
        )
    };
    "Teams" = @{
        Name = "Microsoft Teams"
        AppName = "Office 365 Teams"
        AppRoles = @(
            @{
                Name = "User.Read.All"
                Description = "ユーザー情報へのアクセス"
            },
            @{
                Name = "Directory.Read.All"
                Description = "組織構造情報へのアクセス"
            },
            @{
                Name = "Team.ReadWrite.All"
                Description = "Teams管理"
            },
            @{
                Name = "Channel.ReadWrite.All"
                Description = "チャネル管理"
            },
            @{
                Name = "Chat.ReadWrite.All"
                Description = "チャット管理"
            }
        )
    };
    "EntraID" = @{
        Name = "Microsoft EntraID"
        AppName = "Microsoft Graph"
        AppRoles = @(
            @{
                Name = "User.Read.All"
                Description = "ユーザー情報へのアクセス"
            },
            @{
                Name = "Directory.ReadWrite.All"
                Description = "ディレクトリデータの管理"
            },
            @{
                Name = "Group.ReadWrite.All"
                Description = "グループ管理"
            },
            @{
                Name = "RoleManagement.ReadWrite.Directory"
                Description = "ロール管理"
            }
        )
    };
    "Exchange" = @{
        Name = "Exchange Online"
        AppName = "Office 365 Exchange Online"
        AppRoles = @(
            @{
                Name = "User.Read.All"
                Description = "ユーザー情報へのアクセス"
            },
            @{
                Name = "Directory.Read.All"
                Description = "組織構造情報へのアクセス"
            },
            @{
                Name = "Mail.ReadWrite.All"
                Description = "メール管理"
            },
            @{
                Name = "MailboxSettings.ReadWrite"
                Description = "メールボックス設定管理"
            },
            @{
                Name = "Calendars.ReadWrite.All"
                Description = "カレンダー管理"
            }
        )
    }
}

# App Role取得関数
function Get-AppRoleByName {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]$ServicePrincipal,
        
        [Parameter(Mandatory = $true)]
        [string]$RoleName
    )
    
    try {
        $appRole = $ServicePrincipal.AppRoles | Where-Object { $_.DisplayName -eq $RoleName -or $_.Value -eq $RoleName }
        
        if (-not $appRole) {
            Write-Log "AppRole '$RoleName' が見つかりませんでした" "WARNING"
            return $null
        }
        
        Write-Log "AppRole '$($appRole.DisplayName)' (ID: $($appRole.Id)) を取得しました" "DEBUG" -NoConsole
        return $appRole
    }
    catch {
        Write-Log "AppRole取得中にエラーが発生しました: $_" "ERROR"
        return $null
    }
}

# パーミッション付与関数
function Grant-ApiPermission {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]$ServicePrincipal,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphAppRole]$AppRole,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$TargetUser
    )
    
    try {
        Write-Log "ユーザー '$($TargetUser.DisplayName)' ($($TargetUser.UserPrincipalName)) にパーミッション '$($AppRole.DisplayName)' を付与しています..." "INFO"
        
        # ユーザーの詳細情報を確認
        Write-Log "ユーザーID: $($TargetUser.Id)" "DEBUG" -NoConsole
        Write-Log "サービスプリンシパルID: $($ServicePrincipal.Id)" "DEBUG" -NoConsole
        Write-Log "アプリロールID: $($AppRole.Id)" "DEBUG" -NoConsole
        
        # ユーザー状態チェック
        try {
            $userDetail = Get-MgUser -UserId $TargetUser.Id -ErrorAction Stop
            $userState = if ($userDetail.AccountEnabled) { "有効" } else { "無効" }
            Write-Log "ユーザーアカウントの状態: $userState" "INFO"
            
            if (-not $userDetail.AccountEnabled) {
                Write-Log "ユーザーアカウントが無効化されているため、パーミッションを付与できません" "ERROR"
                Write-Host "エラー: ユーザーアカウントが無効化されています。アカウントを有効化してから再試行してください。" -ForegroundColor Red
                return $false
            }
        }
        catch {
            Write-Log "ユーザー状態の確認中にエラーが発生しました: $_" "WARNING"
        }
        
        # 現在のアプリロール割り当てを確認
        try {
            $existingAssignment = Get-MgUserAppRoleAssignment -UserId $TargetUser.Id -All -ErrorAction Stop | Where-Object { 
                $_.ResourceId -eq $ServicePrincipal.Id -and $_.AppRoleId -eq $AppRole.Id 
            }
            
            if ($existingAssignment) {
                Write-Log "ユーザーには既にこのパーミッションが付与されています" "WARNING"
                Write-Host "情報: ユーザー '$($TargetUser.DisplayName)' には既にパーミッション '$($AppRole.DisplayName)' が付与されています" -ForegroundColor Yellow
                return $false
            }
        }
        catch {
            Write-Log "既存の割り当て確認中にエラーが発生しました: $_" "ERROR"
            Write-Host "エラー: 既存のパーミッション確認中に問題が発生しました。詳細はログファイルを確認してください。" -ForegroundColor Red
            Write-ErrorDetail $_ "既存のパーミッション確認中にエラーが発生"
            return $false
        }
        
        # 新しいアプリロール割り当てを作成
        try {
            $params = @{
                PrincipalId = $TargetUser.Id
                ResourceId = $ServicePrincipal.Id
                AppRoleId = $AppRole.Id
            }
            
            Write-Log "パーミッション割り当てパラメータ: $($params | ConvertTo-Json -Compress)" "DEBUG" -NoConsole
            
            $newAssignment = New-MgUserAppRoleAssignment -UserId $TargetUser.Id -BodyParameter $params -ErrorAction Stop
            
            if ($newAssignment) {
                Write-Log "パーミッション '$($AppRole.DisplayName)' の付与に成功しました" "SUCCESS"
                Write-Host "成功: ユーザー '$($TargetUser.DisplayName)' にパーミッション '$($AppRole.DisplayName)' を付与しました" -ForegroundColor Green
                return $true
            }
            else {
                Write-Log "パーミッション付与に失敗しました（結果が空）" "ERROR"
                Write-Host "エラー: パーミッション付与に失敗しました。操作は完了しましたが、結果が返されませんでした。" -ForegroundColor Red
                return $false
            }
        }
        catch {
            # エラーの種類に応じた詳細メッセージ
            $errorMsg = "詳細な理由: "
            
            if ($_.Exception.Message -match "Permission") {
                $errorMsg += "権限不足。必要な管理者権限があることを確認してください。"
            }
            elseif ($_.Exception.Message -match "AppRole") {
                $errorMsg += "指定されたAPIパーミッション（AppRole）が見つからないか、無効です。"
            }
            elseif ($_.Exception.Message -match "Resource") {
                $errorMsg += "指定されたリソース（ServicePrincipal）が見つからないか、アクセスできません。"
            }
            elseif ($_.Exception.Message -match "Principal") {
                $errorMsg += "指定されたユーザー（Principal）が見つからないか、アクセスできません。"
            }
            elseif ($_.Exception.Message -match "Conflict") {
                $errorMsg += "競合が発生しました。既存の割り当てが存在する可能性があります。"
            }
            elseif ($_.Exception.Message -match "Unauthorized") {
                $errorMsg += "認証エラー。サインインしているアカウントに必要な権限がない可能性があります。"
            }
            else {
                $errorMsg += $_.Exception.Message
            }
            
            Write-Log "パーミッション付与中にエラーが発生: $errorMsg" "ERROR"
            Write-Host "エラー: パーミッション '$($AppRole.DisplayName)' の付与に失敗しました" -ForegroundColor Red
            Write-Host $errorMsg -ForegroundColor Red
            
            # 詳細なエラー情報をログに記録
            Write-ErrorDetail $_ "パーミッション付与中にエラーが発生しました"
            return $false
        }
    }
    catch {
        # 予期しない全体的なエラー
        Write-Log "予期しないエラーが発生しました: $_" "ERROR"
        Write-Host "致命的なエラー: パーミッション処理中に予期しない問題が発生しました。詳細はログを確認してください。" -ForegroundColor Red
        Write-ErrorDetail $_ "パーミッション付与中に予期しないエラーが発生"
        return $false
    }
}

# パーミッション削除関数
function Remove-ApiPermission {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]$ServicePrincipal,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphAppRole]$AppRole,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$TargetUser
    )
    
    try {
        Write-Log "ユーザー '$($TargetUser.DisplayName)' ($($TargetUser.UserPrincipalName)) からパーミッション '$($AppRole.DisplayName)' を削除しています..." "INFO"
        
        # 現在のアプリロール割り当てを確認
        $existingAssignment = Get-MgUserAppRoleAssignment -UserId $TargetUser.Id -All | Where-Object { 
            $_.ResourceId -eq $ServicePrincipal.Id -and $_.AppRoleId -eq $AppRole.Id 
        }
        
        if (-not $existingAssignment) {
            Write-Log "ユーザーにはこのパーミッションが付与されていません" "WARNING"
            return $false
        }
        
        # アプリロール割り当てを削除
        Remove-MgUserAppRoleAssignment -UserId $TargetUser.Id -AppRoleAssignmentId $existingAssignment.Id
        
        Write-Log "パーミッション '$($AppRole.DisplayName)' の削除に成功しました" "SUCCESS"
        return $true
    }
    catch {
        Write-ErrorDetail $_ "パーミッション削除中にエラーが発生しました"
        return $false
    }
}

# すべてのパーミッションを一括付与する関数
function Grant-AllApiPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]$ServicePrincipal,
        
        [Parameter(Mandatory = $true)]
        [array]$AppRoles,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$TargetUser
    )
    
    try {
        Write-Log "ユーザー '$($TargetUser.DisplayName)' ($($TargetUser.UserPrincipalName)) に全パーミッションを一括付与しています..." "INFO"
        
        $successCount = 0
        $failCount = 0
        
        foreach ($roleName in $AppRoles) {
            $appRole = Get-AppRoleByName -ServicePrincipal $ServicePrincipal -RoleName $roleName.Name
            
            if ($appRole) {
                $result = Grant-ApiPermission -ServicePrincipal $ServicePrincipal -AppRole $appRole -TargetUser $TargetUser
                
                if ($result) {
                    $successCount++
                }
                else {
                    $failCount++
                }
            }
            else {
                Write-Log "AppRole '$($roleName.Name)' が見つからないためスキップします" "WARNING"
                $failCount++
            }
        }
        
        Write-Log "一括付与結果: 成功=$successCount, 失敗=$failCount" "INFO"
        return ($failCount -eq 0)
    }
    catch {
        Write-ErrorDetail $_ "一括パーミッション付与中にエラーが発生しました"
        return $false
    }
}

# 現在のパーミッション状態を確認する関数
function Get-CurrentPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphServicePrincipal]$ServicePrincipal,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]$TargetUser
    )
    
    try {
        Write-Log "ユーザー '$($TargetUser.DisplayName)' ($($TargetUser.UserPrincipalName)) の現在のパーミッションを確認しています..." "INFO"
        
        $currentAssignments = Get-MgUserAppRoleAssignment -UserId $TargetUser.Id -All | Where-Object { 
            $_.ResourceId -eq $ServicePrincipal.Id 
        }
        
        if (-not $currentAssignments -or $currentAssignments.Count -eq 0) {
            Write-Log "このアプリケーションに対するパーミッションはありません" "INFO"
            return @()
        }
        
        $permissions = @()
        foreach ($assignment in $currentAssignments) {
            $appRole = $ServicePrincipal.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId }
            
            if ($appRole) {
                $permissions += [PSCustomObject]@{
                    DisplayName = $appRole.DisplayName
                    Value = $appRole.Value
                    Id = $appRole.Id
                    AssignmentId = $assignment.Id
                }
            }
        }
        
        Write-Log "$($permissions.Count) 件のパーミッションが見つかりました" "INFO"
        return $permissions
    }
    catch {
        Write-ErrorDetail $_ "現在のパーミッション確認中にエラーが発生しました"
        return @()
    }
}

# ユーザーの表示（サムアカウント名付き）
function Show-UsersWithSamAccountName {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Users
    )
    
    $usersWithSAM = @()
    foreach ($user in $Users) {
        try {
            # オンプレミスプロパティを取得
            $userDetail = Get-MgUser -UserId $user.Id -Property OnPremisesSamAccountName, DisplayName, UserPrincipalName, Id
            
            $samAccountName = if ($userDetail.AdditionalProperties.onPremisesSamAccountName) {
                $userDetail.AdditionalProperties.onPremisesSamAccountName
            } else {
                "N/A"
            }
            
            $usersWithSAM += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                SamAccountName = $samAccountName
                Id = $user.Id
                User = $user
            }
        }
        catch {
            Write-Log "ユーザー情報取得中にエラー: $($user.UserPrincipalName) - $_" "WARNING"
            $usersWithSAM += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                SamAccountName = "取得エラー"
                Id = $user.Id
                User = $user
            }
        }
    }
    
    # 表形式で表示
    Write-Host "`n===== ユーザー一覧 =====" -ForegroundColor Cyan
    $format = "{0,-3} | {1,-20} | {2,-40} | {3,-15}"
    Write-Host ($format -f "No.", "表示名", "UPN", "SAMアカウント名")
    Write-Host ("-" * 85)
    
    for ($i = 0; $i -lt $usersWithSAM.Count; $i++) {
        # 表示名を処理（長すぎる場合は省略）
        $displayName = $usersWithSAM[$i].DisplayName
        if ($displayName.Length -gt 18) {
            $displayName = $displayName.Substring(0, 15) + "..."
        }
        
        # UPNを処理（長すぎる場合は省略）
        $upn = $usersWithSAM[$i].UserPrincipalName
        if ($upn.Length -gt 38) {
            $upn = $upn.Substring(0, 35) + "..."
        }
        
        # 表示
        Write-Host ($format -f ($i+1), $displayName, $upn, $usersWithSAM[$i].SamAccountName)
    }
    
    return $usersWithSAM
}

# メイン処理の開始点
Write-Log "処理を開始します" "INFO"
try {
    # Microsoft Graph モジュールの確認とインストール
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log "Microsoft Graph モジュールがインストールされていません。インストールを試みます..." "INFO"
        try {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
            Write-Log "Microsoft Graph モジュールのインストールに成功しました" "SUCCESS"
        }
        catch {
            Write-ErrorDetail $_ "Microsoft Graph モジュールのインストールに失敗しました"
            Write-Log "管理者権限で次のコマンドを実行してください: Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force" "ERROR"
            exit 1
        }
    }
    
    # 必要なモジュールをインポート
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Applications
    
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
    
    # 対象システムの選択
    $systemKeys = @($systemDefinitions.Keys)
    $systemOptions = $systemKeys | ForEach-Object { $systemDefinitions[$_].Name }
    $systemChoice = Show-Menu -Title "対象システムの選択" -Options $systemOptions
    if ($systemChoice -eq "Q") {
        Write-Log "ユーザーによりスクリプトが終了されました" "INFO"
        exit 0
    }
    
    $selectedKey = $systemKeys[$systemChoice]
    $selectedSystem = $systemDefinitions[$selectedKey]
    Write-Log "選択されたシステム: $($selectedSystem.Name) (キー: $selectedKey)" "INFO"
    
    # サービスプリンシパル取得
    $servicePrincipal = Get-ServicePrincipalForSystem -SystemInfo $selectedSystem
    if (-not $servicePrincipal) {
        Write-Log "サービスプリンシパル取得に失敗しました。スクリプトを終了します。" "ERROR"
        exit 1
    }
    
    # ユーザー選択
    $targetUsers = Select-TargetUsers -AllowMultiple:($mainChoice -ne 2)
    if (-not $targetUsers -or $targetUsers.Count -eq 0) {
        Write-Log "ユーザーが選択されませんでした。スクリプトを終了します。" "WARNING"
        exit 0
    }
    
    # SAMアカウント名を含めて表示
    $displayedUsers = Show-UsersWithSamAccountName -Users $targetUsers
    
    # 機能に応じた処理
    switch ($mainChoice) {
        0 { # APIパーミッションの付与
            $permissionOptions = @(
                "個別のパーミッションを選択する",
                "システムに最適なパーミッションセットを一括付与する"
            )
            
            $permChoice = Show-Menu -Title "パーミッション付与方法" -Options $permissionOptions
            if ($permChoice -eq "Q") {
                Write-Log "ユーザーによりスクリプトが終了されました" "INFO"
                exit 0
            }
            
            if ($permChoice -eq 0) { # 個別のパーミッションを選択
                $roleOptions = $selectedSystem.AppRoles | ForEach-Object { "$($_.Name) - $($_.Description)" }
                $roleChoice = Show-Menu -Title "付与するAPIパーミッション" -Options $roleOptions
                if ($roleChoice -eq "Q") {
                    Write-Log "ユーザーによりスクリプトが終了されました" "INFO"
                    exit 0
                }
                
                $selectedRole = $selectedSystem.AppRoles[$roleChoice]
                Write-Log "選択されたパーミッション: $($selectedRole.Name)" "INFO"
                
                $appRole = Get-AppRoleByName -ServicePrincipal $servicePrincipal -RoleName $selectedRole.Name
                if (-not $appRole) {
                    Write-Log "AppRoleの取得に失敗しました。スクリプトを終了します。" "ERROR"
                    exit 1
                }
                
                $successCount = 0
                $failCount = 0
                
                foreach ($user in $targetUsers) {
                    $result = Grant-ApiPermission -ServicePrincipal $servicePrincipal -AppRole $appRole -TargetUser $user
                    if ($result) {
                        $successCount++
                    }
                    else {
                        $failCount++
                    }
                }
                
                Write-Log "パーミッション付与結果: 成功=$successCount, 失敗=$failCount" "INFO"
            }
            else { # 一括付与
                $successCount = 0
                $failCount = 0
                
                foreach ($user in $targetUsers) {
                    $result = Grant-AllApiPermissions -ServicePrincipal $servicePrincipal -AppRoles $selectedSystem.AppRoles -TargetUser $user
                    if ($result) {
                        $successCount++
                    }
                    else {
                        $failCount++
                    }
                }
                
                Write-Log "一括パーミッション付与結果: 成功=$successCount, 失敗=$failCount" "INFO"
            }
        }
        1 { # APIパーミッションの削除
            # 現在のパーミッションを表示して選択
            if ($targetUsers.Count -gt 1) {
                Write-Log "パーミッション削除は一度に1ユーザーのみ対応しています。最初のユーザーを処理します。" "WARNING"
            }
            
            $user = $targetUsers[0]
            $currentPermissions = Get-CurrentPermissions -ServicePrincipal $servicePrincipal -TargetUser $user
            
            if ($currentPermissions.Count -eq 0) {
                Write-Log "$($user.DisplayName) には削除するパーミッションがありません" "WARNING"
                exit 0
            }
            
            $permOptions = $currentPermissions | ForEach-Object { "$($_.DisplayName) - $($_.Value)" }
            $permChoice = Show-Menu -Title "削除するAPIパーミッション" -Options $permOptions
            if ($permChoice -eq "Q") {
                Write-Log "ユーザーによりスクリプトが終了されました" "INFO"
                exit 0
            }
            
            $selectedPerm = $currentPermissions[$permChoice]
            
            # AppRoleオブジェクトを作成
            $appRole = [PSCustomObject]@{
                Id = $selectedPerm.Id
                DisplayName = $selectedPerm.DisplayName
            }
            
            $result = Remove-ApiPermission -ServicePrincipal $servicePrincipal -AppRole $appRole -TargetUser $user
            if ($result) {
                Write-Log "パーミッション削除に成功しました" "SUCCESS"
            }
            else {
                Write-Log "パーミッション削除に失敗しました" "ERROR"
            }
        }
        2 { # 現在のパーミッション状態を確認
            $user = $targetUsers[0]
            $currentPermissions = Get-CurrentPermissions -ServicePrincipal $servicePrincipal -TargetUser $user
            
            if ($currentPermissions.Count -eq 0) {
                Write-Log "$($user.DisplayName) にはパーミッションが設定されていません" "INFO"
            }
            else {
                Write-Host "`n===== $($user.DisplayName) の現在のパーミッション =====" -ForegroundColor Cyan
                $format = "{0,-40} | {1,-40}"
                Write-Host ($format -f "パーミッション名", "値")
                Write-Host ("-" * 85)
                
                foreach ($perm in $currentPermissions) {
                    Write-Host ($format -f $perm.DisplayName, $perm.Value)
                }
            }
        }
    }
    
    # 実行サマリーを記録
    Write-ExecutionSummary
    
    # スクリプト終了時に一時停止（プロンプトが閉じないようにする）
    Write-Host "`n処理が完了しました。何かキーを押すと終了します..." -ForegroundColor Green
    if ($Host.Name -eq "ConsoleHost") {
        # コンソールから実行されている場合は、キー入力を待つ
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    else {
        # PowerShell ISEやその他のホストからの実行時は一時停止
        Start-Sleep -Seconds 3
    }
}
catch {
    Write-ErrorDetail $_ "メイン処理の実行中にエラーが発生しました"
    Write-Log "スクリプトは異常終了しました。ログを確認してください: $logFile" "ERROR"
    
    # エラー発生時も一時停止（ユーザーがエラーメッセージを確認できるようにする）
    Write-Host "`nエラーが発生しました。何かキーを押すと終了します..." -ForegroundColor Red
    if ($Host.Name -eq "ConsoleHost") {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    else {
        Start-Sleep -Seconds 5
    }
    
    exit 1
}
