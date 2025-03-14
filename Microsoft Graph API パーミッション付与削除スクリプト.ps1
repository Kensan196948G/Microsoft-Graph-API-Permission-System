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
            $searchQuery = Read-Host "ユーザー名、メールアドレス、またはログイン名(SAMACCOUNTNAME)の一部を入力してください（検索用）"
            
            try {
                Write-Log "ユーザーを検索しています: $searchQuery" "INFO"
                
                # 検索対象の属性リスト
                $searchAttributes = @(
                    @{ Name = "表示名"; Field = "displayName" },
                    @{ Name = "メールアドレス"; Field = "userPrincipalName" },
                    @{ Name = "名"; Field = "givenName" },
                    @{ Name = "姓"; Field = "surname" }
                )
                
                $foundUsers = @()
                
                # API経由で検索可能な属性を検索
                foreach ($attr in $searchAttributes) {
                    try {
                        # Microsoft Graph APIは前方一致検索のみサポート
                        Write-Log "$($attr.Name)で検索中..." "DEBUG" -NoConsole
                        $filter = "startswith($($attr.Field),'$searchQuery')"
                        $users = Get-MgUser -Filter $filter -Top 10 -Property DisplayName, UserPrincipalName, Id, OnPremisesSamAccountName, GivenName, Surname -ErrorAction Stop
                        
                        if ($users -and $users.Count -gt 0) {
                            Write-Log "$($attr.Name)での検索で $($users.Count) 件ヒットしました" "DEBUG" -NoConsole
                            $foundUsers += $users
                        }
                    }
                    catch {
                        Write-Log "$($attr.Name)での検索に失敗しました: $_" "DEBUG" -NoConsole
                        # エラーが発生しても続行
                        continue
                    }
                }
                
                # SAMアカウント名での検索（APIでフィルタリングできないため、クライアント側でフィルタリング）
                try {
                    Write-Log "SAMアカウント名を含む追加検索を実行中..." "DEBUG" -NoConsole
                    
                    # 詳細検索のためにすべてのユーザーを取得（最大100人）
                    # 注意: 大規模な組織では全ユーザーを取得できない可能性がある
                    $allUsers = Get-MgUser -Top 100 -Property DisplayName, UserPrincipalName, Id, OnPremisesSamAccountName, GivenName, Surname -ErrorAction Stop
                    Write-Log "詳細検索のためにユーザーを取得しました（最大100人）" "DEBUG" -NoConsole
                    
                    # クライアント側で複数条件の検索を実行（SAMアカウント名、表示名など）
                    $clientFilteredUsers = $allUsers | Where-Object {
                        # SAMアカウント名での検索
                        ($_.OnPremisesSamAccountName -and (
                            $_.OnPremisesSamAccountName -eq $searchQuery -or                      # 完全一致
                            $_.OnPremisesSamAccountName.StartsWith($searchQuery) -or              # 前方一致
                            $_.OnPremisesSamAccountName.ToLower().Contains($searchQuery.ToLower()) # 部分一致（大文字小文字区別なし）
                        )) -or
                        # 表示名での検索
                        ($_.DisplayName -and (
                            $_.DisplayName -eq $searchQuery -or                      # 完全一致
                            $_.DisplayName.ToLower().Contains($searchQuery.ToLower()) # 部分一致（大文字小文字区別なし）
                        )) -or 
                        # メールアドレスでの検索（UPNの@より前の部分）
                        ($_.UserPrincipalName -and (
                            ($_.UserPrincipalName.Split('@').Length -gt 0 -and
                             $_.UserPrincipalName.Split('@')[0].ToLower().Contains($searchQuery.ToLower())) # ドメイン前の部分で部分一致
                        )) -or
                        # 名前での検索
                        ($_.GivenName -and $_.GivenName.ToLower().Contains($searchQuery.ToLower())) -or
                        # 姓での検索
                        ($_.Surname -and $_.Surname.ToLower().Contains($searchQuery.ToLower()))
                    }
                    
                    if ($clientFilteredUsers -and $clientFilteredUsers.Count -gt 0) {
                        Write-Log "クライアント側の詳細検索で $($clientFilteredUsers.Count) 件ヒットしました" "DEBUG" -NoConsole
                        $foundUsers += $clientFilteredUsers
                    }
                }
                catch {
                    Write-Log "SAMアカウント名での検索に失敗しました: $_" "DEBUG" -NoConsole
                    # エラーが発生しても続行
                }
                
                # 重複を削除
                $foundUsers = $foundUsers | Sort-Object -Property Id -Unique |
                             Select-Object DisplayName, UserPrincipalName, Id, OnPremisesSamAccountName, GivenName, Surname
                
                if ($foundUsers.Count -eq 0) {
                    Write-Log "検索条件「$searchQuery」に一致するユーザーが見つかりませんでした" "WARNING"
                    return $null
                }
                
                # 見つかったユーザー数の変数
                $userCount = $foundUsers.Count
                
                # 完全に別の変数として作成
                $countMessage = [string]::Format("{0} 人のユーザーが見つかりました", $userCount)
                Write-Log $countMessage "INFO"
                
                # ユーザーが1人の場合は、後続処理を一切スキップして即時リターン
                if ($userCount -eq 1) {
                    $singleUser = $foundUsers[0]
                    $userName = $singleUser.DisplayName
                    Write-Log "検索結果が1件のみのため、自動的に選択します: $userName" "INFO"
                    return @($singleUser)
                }
                
                # より詳細な情報を表示するオプションリストを作成
                $userOptions = $foundUsers | ForEach-Object {
                    $details = "$($_.DisplayName) ($($_.UserPrincipalName))"
                    
                    # SAMアカウント名があれば追加
                    if ($_.OnPremisesSamAccountName) {
                        $details += " [SAM: $($_.OnPremisesSamAccountName)]"
                    }
                    
                    # 姓名の情報を追加（表示名と異なる場合のみ）
                    if ($_.GivenName -and $_.Surname -and "$($_.Surname) $($_.GivenName)" -ne $_.DisplayName) {
                        $details += " - $($_.GivenName) $($_.Surname)"
                    }
                    
                    return $details
                }
                # ユーザーオプションの準備が完了したら、複数件ある場合の選択処理
                if ($userCount -gt 1) {
                    if ($AllowMultiple) {
                        # 複数ユーザー選択
                        Write-Host "`n複数のユーザーを選択できます。選択を終了するには 'done' と入力してください。"
                        
                        for ($i = 0; $i -lt $userOptions.Count; $i++) {
                            Write-Host "$($i+1). $($userOptions[$i])"
                        }
                        
                        do {
                            $choice = Read-Host "ユーザー番号を入力 (複数可、カンマ区切り。終了は 'done')"
                            
                            if ($choice -eq "done") { break }
                            
                            $choices = $choice -split "," | ForEach-Object { $_.Trim() }
                            
                            foreach ($c in $choices) {
                                if ([int]::TryParse($c, [ref]$null) -and [int]$c -ge 1 -and [int]$c -le $userOptions.Count) {
                                    $index = [int]$c - 1
                                    if (-not ($selectedUsers -contains $foundUsers[$index])) {
                                        $selectedUsers += $foundUsers[$index]
                                        Write-Host "  + 追加: $($foundUsers[$index].DisplayName)" -ForegroundColor Cyan
                                    }
                                }
                            }
                            
                            Write-Host "  現在の選択ユーザー数: $($selectedUsers.Count)" -ForegroundColor Yellow
                        } while ($true)
                    }
                    else {
                        # 単一ユーザー選択
                        $userChoice = Show-Menu -Title "ユーザーを選択" -Options $userOptions
                        if ($userChoice -eq "Q") { return $null }
                        $selectedUsers = @($foundUsers[$userChoice])
                    }
                }
            }
            catch {
                Write-Log "ユーザー検索中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
        1 { # セキュリティグループからユーザーを選択
            try {
                Write-Log "セキュリティグループを取得しています..." "INFO"
                $groups = Get-MgGroup -Filter "securityEnabled eq true" -Top 20 -ErrorAction Stop | 
                          Select-Object DisplayName, Description, Id
                
                if ($groups.Count -eq 0) {
                    Write-Log "セキュリティグループが見つかりませんでした" "WARNING"
                    return $null
                }
                
                $groupOptions = $groups | ForEach-Object { "$($_.DisplayName) - $($_.Description)" }
                $groupChoice = Show-Menu -Title "セキュリティグループを選択" -Options $groupOptions
                
                if ($groupChoice -eq "Q") { return $null }
                
                $selectedGroup = $groups[$groupChoice]
                Write-Log "グループ「$($selectedGroup.DisplayName)」のメンバーを取得しています..." "INFO"
                
                # グループメンバーの取得
                $groupMembers = Get-MgGroupMember -GroupId $selectedGroup.Id -ErrorAction Stop
                
                if ($groupMembers.Count -eq 0) {
                    Write-Log "選択したグループにメンバーがいません" "WARNING"
                    return $null
                }
                
                # ユーザーの詳細情報を取得
                $selectedUsers = @()
                foreach ($member in $groupMembers) {
                    try {
                        $user = Get-MgUser -UserId $member.Id -ErrorAction SilentlyContinue
                        if ($user) {
                            $selectedUsers += $user | Select-Object DisplayName, UserPrincipalName, Id
                        }
                    }
                    catch {
                        # ユーザーでないメンバーはスキップ
                        continue
                    }
                }
                
                Write-Log "$($selectedUsers.Count) 人のユーザーが選択されました" "INFO"
                
                # 選択されたユーザーの確認
                Write-Host "`n選択されたユーザー:"
                $selectedUsers | ForEach-Object { Write-Host "  * $($_.DisplayName) ($($_.UserPrincipalName))" }
                
                $confirm = Read-Host "これらのユーザーに対してパーミッション操作を行いますか？(Y/N)"
                if ($confirm -ne "Y" -and $confirm -ne "y") {
                    return $null
                }
            }
            catch {
                Write-Log "グループ操作中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
        2 { # CSVファイルからユーザーをインポート
            $csvPath = Read-Host "CSVファイルのパスを入力してください（UserPrincipalName列必須）"
            
            if (-not (Test-Path $csvPath)) {
                Write-Log "CSVファイルが見つかりません: $csvPath" "ERROR"
                return $null
            }
            
            try {
                Write-Log "CSVファイルからユーザーデータをインポートしています..." "INFO"
                $csvUsers = Import-Csv -Path $csvPath -ErrorAction Stop
                
                if (-not ($csvUsers | Get-Member -Name "UserPrincipalName")) {
                    Write-Log "CSVファイルに必須の'UserPrincipalName'列がありません" "ERROR"
                    return $null
                }
                
                $selectedUsers = @()
                $notFoundUsers = @()
                
                foreach ($csvUser in $csvUsers) {
                    try {
                        $user = Get-MgUser -Filter "userPrincipalName eq '$($csvUser.UserPrincipalName)'" -ErrorAction Stop
                        if ($user) {
                            $selectedUsers += $user | Select-Object DisplayName, UserPrincipalName, Id
                        }
                        else {
                            $notFoundUsers += $csvUser.UserPrincipalName
                        }
                    }
                    catch {
                        $notFoundUsers += $csvUser.UserPrincipalName
                        Write-Log "ユーザー取得中にエラー: $($csvUser.UserPrincipalName) - $_" "WARNING"
                    }
                }
                
                if ($notFoundUsers.Count -gt 0) {
                    Write-Log "$($notFoundUsers.Count) 人のユーザーが見つかりませんでした" "WARNING"
                    Write-Host "見つからなかったユーザー:" -ForegroundColor Yellow
                    $notFoundUsers | ForEach-Object { Write-Host "  * $_" -ForegroundColor Yellow }
                }
                
                Write-Log "$($selectedUsers.Count) 人のユーザーがCSVからインポートされました" "INFO"
                
                # 選択されたユーザーの確認
                Write-Host "`n選択されたユーザー:"
                $selectedUsers | ForEach-Object { Write-Host "  * $($_.DisplayName) ($($_.UserPrincipalName))" }
                
                $confirm = Read-Host "これらのユーザーに対してパーミッション操作を行いますか？(Y/N)"
                if ($confirm -ne "Y" -and $confirm -ne "y") {
                    return $null
                }
            }
            catch {
                Write-Log "CSVインポート中にエラーが発生しました: $_" "ERROR"
                return $null
            }
        }
    }
    
    return $selectedUsers
}

# 実行ポリシーの設定（管理者権限で実行）
try {
    Write-Log "スクリプトを開始しています..." "INFO"
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop
    Write-Log "実行ポリシーが設定されました（RemoteSigned）" "INFO"
}
catch [System.UnauthorizedAccessException] {
    Write-Log "管理者権限がないため実行ポリシーを変更できませんでした。一部の機能が制限される可能性があります。" "WARNING"
    # エラーを無視して続行
    Write-Host "実行ポリシーの変更ができませんでしたが、処理を続行します..." -ForegroundColor Yellow
}
catch {
    Write-Log "実行ポリシーの設定中にエラーが発生しました: $_" "ERROR"
    exit 1
}

# Microsoft Graph PowerShell モジュールのインストール
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    try {
        Write-Log "Microsoft Graph モジュールをインストールしています..." "INFO"
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "Microsoft Graph モジュールのインストールに成功しました" "SUCCESS"
    }
    catch [System.UnauthorizedAccessException] {
        try {
            Write-Log "管理者権限がないため、現在のユーザースコープでインストールを試みます..." "WARNING"
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Microsoft Graph モジュールのインストールに成功しました（CurrentUserスコープ）" "SUCCESS"
        }
        catch {
            Write-Log "Microsoft Graph モジュールのインストールに失敗しました: $_" "ERROR"
            Write-Host "Microsoft Graph モジュールのインストールに失敗しました。" -ForegroundColor Red
            Write-Host "以下のエラーが発生しました：" -ForegroundColor Red
            Write-Host $_
            exit 1
        }
    }
    catch {
        Write-Log "Microsoft Graph モジュールのインストールに失敗しました: $_" "ERROR"
        Write-Host "Microsoft Graph モジュールのインストールに失敗しました。" -ForegroundColor Red
        Write-Host "以下のエラーが発生しました：" -ForegroundColor Red
        Write-Host $_
        exit 1
    }
}

# モジュールをインポート
try {
    Write-Log "Microsoft Graph モジュールをインポートしています..." "INFO"
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Applications
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Groups
    Write-Log "モジュールのインポートに成功しました" "SUCCESS"
}
catch {
    Write-Log "モジュールのインポート中にエラーが発生しました: $_" "ERROR"
    exit 1
}

# Microsoft Graphに接続
if (-not (Connect-ToGraph)) {
    Write-Log "Microsoft Graphへの接続に失敗しました。スクリプトを終了します。" "ERROR"
    exit 1
}

# 管理者権限を確認
if (-not (Test-AdminRole)) {
    Write-Log "必要な管理者権限がありません。スクリプトを終了します。" "ERROR"
    exit 1
}

# 対象のシステムを選択（最適なパーミッションセットを定義）
$systems = @(
    @{ 
        Name = "OneDrive for Business"; 
        AppName = "Microsoft Graph"; # Office 365 SharePoint OnlineからMicrosoft Graphに変更
        OptimalPermissions = @(
            @{ Name = "User.Read.All"; Description = "ユーザー情報へのアクセス"; AlternativeNames = @("User.Read", "User.ReadBasic.All") },
            @{ Name = "Directory.Read.All"; Description = "組織構造情報へのアクセス"; AlternativeNames = @("Directory.AccessAsUser.All", "Organization.Read.All") },
            @{ Name = "Files.ReadWrite.All"; Description = "OneDriveファイルの読み書き"; AlternativeNames = @("Sites.ReadWrite.All", "AllSites.Write") }
        )
    },
    @{ 
        Name = "Microsoft Teams"; 
        AppName = "Microsoft Teams Services";
        OptimalPermissions = @(
            @{ Name = "User.Read.All"; Description = "ユーザー情報へのアクセス"; AlternativeNames = @("User.Read", "User.ReadBasic.All") },
            @{ Name = "Directory.Read.All"; Description = "組織構造情報へのアクセス"; AlternativeNames = @("Directory.AccessAsUser.All", "Organization.Read.All") },
            @{ Name = "Team.ReadWrite.All"; Description = "Teams管理"; AlternativeNames = @() },
            @{ Name = "Channel.ReadWrite.All"; Description = "チャネル管理"; AlternativeNames = @() },
            @{ Name = "Chat.ReadWrite.All"; Description = "チャット管理"; AlternativeNames = @() }
        )
    },
    @{ 
        Name = "Microsoft EntraID"; 
        AppName = "Microsoft Graph";
        OptimalPermissions = @(
            @{ Name = "User.Read.All"; Description = "ユーザー情報へのアクセス" },
            @{ Name = "Directory.ReadWrite.All"; Description = "ディレクトリデータの管理" },
            @{ Name = "Group.ReadWrite.All"; Description = "グループ管理" },
            @{ Name = "RoleManagement.ReadWrite.Directory"; Description = "ロール管理" }
        )
    },
    @{ 
        Name = "Exchange Online"; 
        AppName = "Office 365 Exchange Online";
        OptimalPermissions = @(
            @{ Name = "User.Read.All"; Description = "ユーザー情報へのアクセス" },
            @{ Name = "Directory.Read.All"; Description = "組織構造情報へのアクセス" },
            @{ Name = "Mail.ReadWrite.All"; Description = "メール管理" },
            @{ Name = "MailboxSettings.ReadWrite"; Description = "メールボックス設定管理" },
            @{ Name = "Calendars.ReadWrite.All"; Description = "カレンダー管理" }
        )
    }
)

$systemOptions = $systems | ForEach-Object { $_.Name }
$systemChoice = Show-Menu -Title "対象のシステムを選択" -Options $systemOptions

if ($systemChoice -eq "Q") {
    Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
    exit 0
}

$selectedSystem = $systems[$systemChoice]

# 選択したアプリの Service Principal ID を取得
$servicePrincipal = Get-ServicePrincipalForSystem -SystemInfo $selectedSystem

if (-not $servicePrincipal) {
    Write-Log "対象のアプリを見つけることができませんでした。スクリプトを終了します。" "ERROR"
    exit 1
}

# アクションの選択
$actionOptions = @(
    "最適なパーミッションセットを一括付与",
    "個別にパーミッションを付与",
    "パーミッションを削除",
    "現在のパーミッション割り当てを表示"
)
$actionChoice = Show-Menu -Title "実行するアクション" -Options $actionOptions

if ($actionChoice -eq "Q") {
    Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
    exit 0
}

# 現在のパーミッション割り当てを表示する場合
if ($actionChoice -eq 3) {
    try {
        Write-Log "現在のパーミッション割り当てを取得しています..." "INFO"
        
        # アプリロールの取得
        $appRoles = $servicePrincipal.AppRoles | Where-Object { $_.IsEnabled -eq $true } | 
                    Select-Object DisplayName, Id, Description
        
        # パーミッション割り当ての取得
        $assignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All -ErrorAction Stop
        
        if ($assignments.Count -eq 0) {
            Write-Host "`n現在、「$($selectedSystem.Name)」に対してパーミッションが割り当てられているユーザーはいません。" -ForegroundColor Yellow
        }
        else {
            Write-Host "`n「$($selectedSystem.Name)」の現在のパーミッション割り当て:" -ForegroundColor Cyan
            
            $groupedAssignments = $assignments | Group-Object -Property AppRoleId
            
            foreach ($group in $groupedAssignments) {
                $roleName = ($appRoles | Where-Object { $_.Id -eq $group.Name }).DisplayName
                if (-not $roleName) { $roleName = "不明なロール" }
                
                Write-Host "`n▶ $roleName (ID: $($group.Name))" -ForegroundColor Yellow
                
                foreach ($assignment in $group.Group) {
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($user) {
                            Write-Host "  - $($user.DisplayName) ($($user.UserPrincipalName))"
                        }
                        else {
                            # ユーザー以外の場合（グループなど）
                            $otherPrincipal = Get-MgDirectoryObject -DirectoryObjectId $assignment.PrincipalId -ErrorAction SilentlyContinue
                            Write-Host "  - ID: $($assignment.PrincipalId) [非ユーザープリンシパル]"
                        }
                    }
                    catch {
                        Write-Host "  - ID: $($assignment.PrincipalId) [詳細不明]"
                    }
                }
            }
        }
        
        # ログに記録
        Write-Log "$($assignments.Count) 件のパーミッション割り当てが表示されました" "INFO"
        
        # スクリプト終了
        Write-Host "`n処理が完了しました。Enterキーを押して終了してください..." -ForegroundColor Cyan
        Read-Host
        exit 0
    }
    catch {
        Write-Log "パーミッション割り当ての取得中にエラーが発生しました: $_" "ERROR"
        exit 1
    }
}

# 一括パーミッション付与モード
if ($actionChoice -eq 0) {
    try {
        Write-Log "最適なパーミッションセットを一括付与モードを開始します" "INFO"
        
        # 選択されたシステムの最適なパーミッションセットを取得
        $optimalPermissions = $selectedSystem.OptimalPermissions
        
        if (-not $optimalPermissions -or $optimalPermissions.Count -eq 0) {
            Write-Log "選択されたシステムの最適なパーミッションセットが定義されていません" "ERROR"
            exit 1
        }
        
        Write-Log "$($selectedSystem.Name)の最適なパーミッションセット($($optimalPermissions.Count)個)を適用します" "INFO"
        
        # 全ての利用可能なパーミッションを取得
        $appRoles = $servicePrincipal.AppRoles | Where-Object { $_.IsEnabled -eq $true } |
                    Select-Object DisplayName, Id, Description, Value
        
        if ($appRoles.Count -eq 0) {
            Write-Log "利用可能なAPIパーミッションが見つかりませんでした" "ERROR"
            exit 1
        }
        
        # パーミッションセットを表示
        Write-Host "`n$($selectedSystem.Name)の最適なパーミッションセット:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $optimalPermissions.Count; $i++) {
            Write-Host "$($i+1). $($optimalPermissions[$i].Name) - $($optimalPermissions[$i].Description)" -ForegroundColor White
        }
        
        # 該当するパーミッションIDを見つける
        $selectedRoles = @()
        $notFoundPermissions = @()
        
        foreach ($permission in $optimalPermissions) {
            $found = $false
            
            # Value（パーミッション名）でマッチングを試みる
            $role = $appRoles | Where-Object { $_.Value -eq $permission.Name }
            
            # 代替名でマッチングを試みる
            if (-not $role -and $permission.PSObject.Properties.Name -contains "AlternativeNames") {
                foreach ($altName in $permission.AlternativeNames) {
                    $role = $appRoles | Where-Object { $_.Value -eq $altName }
                    if ($role) { 
                        Write-Log "代替パーミッション名「$altName」でマッチしました" "DEBUG" -NoConsole
                        break 
                    }
                }
            }
            
            # DisplayName（表示名）でマッチングを試みる
            if (-not $role) {
                $role = $appRoles | Where-Object { $_.DisplayName -eq $permission.Name }
            }
            
            # 部分一致でマッチングを試みる（最終手段）
            if (-not $role) {
                $role = $appRoles | Where-Object { 
                    $_.Value -like "*$($permission.Name)*" -or 
                    $_.DisplayName -like "*$($permission.Name)*" 
                } | Select-Object -First 1
                
                if ($role) {
                    Write-Log "部分一致でパーミッション「$($permission.Name)」を「$($role.Value)」にマッピングしました" "DEBUG" -NoConsole
                }
            }
            
            if ($role) {
                $selectedRoles += $role
                $found = $true
            } else {
                $notFoundPermissions += $permission.Name
            }
        }
        
        # 見つからなかったパーミッションがあれば警告
        if ($notFoundPermissions.Count -gt 0) {
            Write-Log "警告: 以下のパーミッションが見つかりませんでした: $($notFoundPermissions -join ', ')" "WARNING"
            Write-Host "`n以下のパーミッションは見つからなかったため、適用されません:" -ForegroundColor Yellow
            foreach ($notFound in $notFoundPermissions) {
                Write-Host "  - $notFound" -ForegroundColor Yellow
            }
            
            $continue = Read-Host "`n見つかったパーミッションのみで続行しますか？ (Y/N)"
            if ($continue -ne "Y" -and $continue -ne "y") {
                Write-Log "ユーザーによって処理がキャンセルされました" "INFO"
                exit 0
            }
        }
        
        if ($selectedRoles.Count -eq 0) {
            Write-Log "適用できるパーミッションが見つかりませんでした" "ERROR"
            exit 1
        }
        
        Write-Log "$($selectedRoles.Count)個のパーミッションを適用します" "INFO"
        
        # 以降は複数ユーザー選択と付与処理に続く
    }
    catch {
        Write-Log "最適なパーミッションセットの処理中にエラーが発生しました: $_" "ERROR"
        exit 1
    }
}
# 個別パーミッション付与モード
elseif ($actionChoice -eq 1) {
    try {
        $appRoles = $servicePrincipal.AppRoles | Where-Object { $_.IsEnabled -eq $true } |
                    Select-Object DisplayName, Id, Description, Value

        if ($appRoles.Count -eq 0) {
            Write-Log "利用可能なAPIパーミッションが見つかりませんでした" "ERROR"
            exit 1
        }

        Write-Log "$($appRoles.Count) 個のAPIパーミッションが見つかりました" "INFO"

        # パーミッション表示の簡略化
        Write-Host "`n利用可能な API パーミッション:" -ForegroundColor Cyan
        $appRoleOptions = @()

        # 詳細情報フラグ
        $showDetailedInfo = $false
        if ($detailedLogEnabled) {
            $showDetailInfo = Read-Host "パーミッションの詳細情報も表示しますか？(Y/N)"
            $showDetailedInfo = ($showDetailInfo -eq "Y" -or $showDetailInfo -eq "y")
        }

        # SharePointフィルタリング
        $showSharePointPermissions = $true
        if ($selectedSystem.Name -eq "OneDrive for Business") {
            $showSharePointFilter = Read-Host "SharePointパーミッションを表示しますか？(Y/N)"
            $showSharePointPermissions = ($showSharePointFilter -eq "Y" -or $showSharePointFilter -eq "y")
        }

        # パーミッションをフィルタリング
        $filteredAppRoles = $appRoles
        if (-not $showSharePointPermissions) {
            # SharePoint関連のパーミッションをフィルタリング
            $filteredAppRoles = $appRoles | Where-Object {
                -not ($_.DisplayName -like "*サイトコレクション*" -or
                     $_.DisplayName -like "*SharePoint*" -or
                     $_.Description -like "*SharePoint*" -or
                     $_.DisplayName -like "*Sites*")
            }
            Write-Log "SharePoint関連のパーミッションを除外しました" "INFO"
        }

        # シンプルに表示するための処理
        for ($i = 0; $i -lt $filteredAppRoles.Count; $i++) {
            # 簡潔な表示用の説明文作成
            $simplifiedDesc = $filteredAppRoles[$i].Description
            if ($simplifiedDesc.Length -gt 60) {
                $simplifiedDesc = $simplifiedDesc.Substring(0, 60) + "..."
            }
            
            $appRoleOptions += "$($filteredAppRoles[$i].DisplayName) - $simplifiedDesc"
            
            # 番号と名前を表示（常に表示）
            Write-Host "$($i+1). $($filteredAppRoles[$i].DisplayName)" -ForegroundColor White
            
            # 詳細情報は必要に応じて表示
            if ($showDetailedInfo) {
                Write-Host "   ID: $($filteredAppRoles[$i].Id)" -ForegroundColor Gray
                Write-Host "   説明: $($filteredAppRoles[$i].Description)" -ForegroundColor Gray
                Write-Host "   値: $($filteredAppRoles[$i].Value)" -ForegroundColor Gray
            }
            
            # 空行を入れる（詳細表示時のみ）
            if ($showDetailedInfo) {
                Write-Host ""
            }
        }

        $roleChoice = Show-Menu -Title "パーミッションを選択" -Options $appRoleOptions

        if ($roleChoice -eq "Q") {
            Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
            exit 0
        }

        $selectedRoles = @($filteredAppRoles[$roleChoice])
        Write-Log "パーミッション「$($selectedRoles[0].DisplayName)」が選択されました" "INFO"
    }
    catch {
        Write-Log "APIパーミッションの取得中にエラーが発生しました: $_" "ERROR"
        exit 1
    }
}
# パーミッション削除モード
elseif ($actionChoice -eq 2) {
    try {
        $appRoles = $servicePrincipal.AppRoles | Where-Object { $_.IsEnabled -eq $true } |
                    Select-Object DisplayName, Id, Description, Value

        if ($appRoles.Count -eq 0) {
            Write-Log "利用可能なAPIパーミッションが見つかりませんでした" "ERROR"
            exit 1
        }

        # パーミッション表示
        Write-Host "`n削除可能な API パーミッション:" -ForegroundColor Cyan
        $appRoleOptions = @()

        for ($i = 0; $i -lt $appRoles.Count; $i++) {
            $simplifiedDesc = $appRoles[$i].Description
            if ($simplifiedDesc.Length -gt 60) {
                $simplifiedDesc = $simplifiedDesc.Substring(0, 60) + "..."
            }
            
            $appRoleOptions += "$($appRoles[$i].DisplayName) - $simplifiedDesc"
            Write-Host "$($i+1). $($appRoles[$i].DisplayName)" -ForegroundColor White
        }

        $roleChoice = Show-Menu -Title "削除するパーミッションを選択" -Options $appRoleOptions

        if ($roleChoice -eq "Q") {
            Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
            exit 0
        }

        $selectedRoles = @($appRoles[$roleChoice])
        Write-Log "削除するパーミッション「$($selectedRoles[0].DisplayName)」が選択されました" "INFO"
    }
    catch {
        Write-Log "APIパーミッションの取得中にエラーが発生しました: $_" "ERROR"
        exit 1
    }
}

# ユーザー選択はアクションの種類に関わらず実行
if ($actionChoice -ne 3) {  # 表示モード以外
    # 複数ユーザー選択を許可するオプション
    $allowMultiple = $true

    # ターゲットユーザーの選択
    $targetUsers = Select-TargetUsers -AllowMultiple:$allowMultiple

    if (-not $targetUsers -or $targetUsers.Count -eq 0) {
        Write-Log "ユーザーが選択されていないか、選択に失敗しました。スクリプトを終了します。" "WARNING"
        exit 0
    }

    Write-Log "$($targetUsers.Count) 人のユーザーが選択されました" "INFO"

    # 処理の概要を表示
    Write-Host "`n以下の操作を実行します：" -ForegroundColor Cyan
    Write-Host "  システム: $($selectedSystem.Name)" -ForegroundColor White
    
    if ($actionChoice -eq 0) {
        Write-Host "  パーミッション: 最適なパーミッションセット($($selectedRoles.Count)個)" -ForegroundColor White
        foreach ($role in $selectedRoles) {
            Write-Host "    - $($role.DisplayName)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  パーミッション: $($selectedRoles[0].DisplayName)" -ForegroundColor White
    }
    
    Write-Host "  アクション: $($actionOptions[$actionChoice])" -ForegroundColor White
    Write-Host "  対象ユーザー数: $($targetUsers.Count)" -ForegroundColor White

    $confirmation = Read-Host "`n処理を続行しますか？(Y/N)"

    if ($confirmation -ne "Y" -and $confirmation -ne "y") {
        Write-Log "ユーザーにより処理がキャンセルされました" "INFO"
        exit 0
    }

    # バッチ処理の開始
    $successCount = 0
    $failureCount = 0
    $skippedCount = 0
    $errorDetails = @()
    $operationLog = @()

    Write-Log "バッチ処理を開始します。対象ユーザー数: $($targetUsers.Count)人" "INFO"
    
    if ($actionChoice -eq 0) {
        Write-Log "処理内容: $($actionOptions[$actionChoice]) - $($selectedRoles.Count)個のパーミッション" "INFO"
    } else {
        Write-Log "処理内容: $($actionOptions[$actionChoice]) - $($selectedRoles[0].DisplayName)" "INFO"
    }

    foreach ($user in $targetUsers) {
        # 複数パーミッションのケース（一括付与モード）
        if ($actionChoice -eq 0) {
            $userSuccessCount = 0
            $userSkippedCount = 0
            $userFailureCount = 0
            
            Write-Log "ユーザー「$($user.DisplayName)」に対して処理を開始します..." "INFO"
            
            foreach ($role in $selectedRoles) {
                $currentRoleLog = @{
                    UserDisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    UserId = $user.Id
                    Action = $actionOptions[$actionChoice]
                    Permission = $role.DisplayName
                    Status = "処理中"
                    Timestamp = Get-Date
                    ErrorMessage = ""
                }
                
                try {
                    Write-Log "  パーミッション「$($role.DisplayName)」を付与しています..." "INFO"
                    
                    # 既存の割り当てを確認
                    $existingAssignment = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All |
                                          Where-Object { $_.PrincipalId -eq $user.Id -and $_.AppRoleId -eq $role.Id }
                    
                    if ($existingAssignment) {
                        Write-Log "  ユーザー「$($user.DisplayName)」には既にパーミッション「$($role.DisplayName)」が付与されています" "WARNING"
                        $currentRoleLog.Status = "スキップ"
                        $currentRoleLog.ErrorMessage = "既に権限が付与済み"
                        $userSkippedCount++
                        $skippedCount++
                        continue
                    }
                    
                    # API呼び出し時間を記録
                    $apiStartTime = Get-Date
                    
                    # パーミッション付与
                    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id `
                        -PrincipalId $user.Id `
                        -ResourceId $servicePrincipal.Id `
                        -AppRoleId $role.Id -ErrorAction Stop
                    
                    # API呼び出し時間を計測
                    $apiDuration = (Get-Date) - $apiStartTime
                    Write-Log "  API呼び出し時間: $($apiDuration.TotalMilliseconds)ms" "VERBOSE" -NoConsole
                        
                    $userSuccessCount++
                    $successCount++
                    $currentRoleLog.Status = "成功"
                    Write-Log "  パーミッション「$($role.DisplayName)」の付与に成功しました" "SUCCESS"
                }
                catch {
                    $userFailureCount++
                    $failureCount++
                    $currentRoleLog.Status = "失敗"
                    $currentRoleLog.ErrorMessage = $_.Exception.Message
                    
                    # 詳細なエラー情報を記録
                    Write-ErrorDetail $_ "ユーザー「$($user.DisplayName)」のパーミッション「$($role.DisplayName)」付与中にエラー"
                    
                    $errorDetails += "$($user.UserPrincipalName) - $($role.DisplayName): $($_.Exception.Message)"
                }
                
                # 操作ログに追加
                $operationLog += $currentRoleLog
            }
            
            # ユーザーごとのサマリー
            Write-Log "ユーザー「$($user.DisplayName)」の処理完了。成功: $userSuccessCount, スキップ: $userSkippedCount, 失敗: $userFailureCount" "INFO"
        }
        # 単一パーミッションのケース（個別付与または削除モード）
        else {
            $currentUserLog = @{
                UserDisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                UserId = $user.Id
                Action = $actionOptions[$actionChoice]
                Permission = $selectedRoles[0].DisplayName
                Status = "処理中"
                Timestamp = Get-Date
                ErrorMessage = ""
            }
            
            try {
                if ($actionChoice -eq 1) {  # 個別パーミッション付与
                    Write-Log "ユーザー「$($user.DisplayName)」にパーミッション「$($selectedRoles[0].DisplayName)」を付与しています..." "INFO"
                    
                    # 既存の割り当てを確認
                    $existingAssignment = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All |
                                          Where-Object { $_.PrincipalId -eq $user.Id -and $_.AppRoleId -eq $selectedRoles[0].Id }
                    
                    if ($existingAssignment) {
                        Write-Log "ユーザー「$($user.DisplayName)」には既にこのパーミッションが付与されています" "WARNING"
                        $currentUserLog.Status = "スキップ"
                        $currentUserLog.ErrorMessage = "既に権限が付与済み"
                        $skippedCount++
                        continue
                    }
                    
                    # API呼び出し時間を記録
                    $apiStartTime = Get-Date
                    
                    # パーミッション付与
                    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id `
                        -PrincipalId $user.Id `
                        -ResourceId $servicePrincipal.Id `
                        -AppRoleId $selectedRoles[0].Id -ErrorAction Stop
                    
                    # API呼び出し時間を計測
                    $apiDuration = (Get-Date) - $apiStartTime
                    Write-Log "API呼び出し時間: $($apiDuration.TotalMilliseconds)ms" "VERBOSE" -NoConsole
                        
                    $successCount++
                    $currentUserLog.Status = "成功"
                    Write-Log "ユーザー「$($user.DisplayName)」へのパーミッション付与に成功しました" "SUCCESS"
                }
                elseif ($actionChoice -eq 2) {  # パーミッション削除
                    Write-Log "ユーザー「$($user.DisplayName)」からパーミッション「$($selectedRoles[0].DisplayName)」を削除しています..." "INFO"
                    
                    # 既存の割り当てを確認
                    $existingAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All |
                                         Where-Object { $_.PrincipalId -eq $user.Id -and $_.AppRoleId -eq $selectedRoles[0].Id }
                    
                    if (-not $existingAssignments -or $existingAssignments.Count -eq 0) {
                        Write-Log "ユーザー「$($user.DisplayName)」にはこのパーミッションが付与されていません" "WARNING"
                        $currentUserLog.Status = "スキップ"
                        $currentUserLog.ErrorMessage = "権限が付与されていない"
                        $skippedCount++
                        continue
                    }
                    
                    # API呼び出し時間を記録
                    $apiStartTime = Get-Date
                    
                    # 各割り当ての削除
                    foreach ($assignment in $existingAssignments) {
                        # パーミッション削除（AppRoleAssignmentId で識別）
                        Remove-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id `
                            -AppRoleAssignmentId $assignment.Id -ErrorAction Stop
                    }
                    
                    # API呼び出し時間を計測
                    $apiDuration = (Get-Date) - $apiStartTime
                    Write-Log "API呼び出し時間: $($apiDuration.TotalMilliseconds)ms" "VERBOSE" -NoConsole
                        
                    $successCount++
                    $currentUserLog.Status = "成功"
                    Write-Log "ユーザー「$($user.DisplayName)」からのパーミッション削除に成功しました" "SUCCESS"
                }
            }
            catch {
                $failureCount++
                $currentUserLog.Status = "失敗"
                $currentUserLog.ErrorMessage = $_.Exception.Message
                
                # 詳細なエラー情報を記録
                Write-ErrorDetail $_ "ユーザー「$($user.DisplayName)」の処理中にエラーが発生しました"
                
                $errorDetails += "$($user.UserPrincipalName): $($_.Exception.Message)"
            }
            
            # 操作ログに追加
            $operationLog += $currentUserLog
        }
    }

# この部分は削除 - 既に上部の修正部分に含まれているため不要

# 詳細な操作ログをファイルに記録
$operationLogContent = "==== 操作詳細ログ ====`n"
$operationLogContent += "処理時刻: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
$operationLogContent += "システム: $($selectedSystem.Name)`n"
$operationLogContent += "パーミッション: $($selectedRole.DisplayName)`n"
$operationLogContent += "アクション: $($actionOptions[$actionChoice])`n"
$operationLogContent += "`n--- ユーザー別処理結果 ---`n"

foreach ($entry in $operationLog) {
    $operationLogContent += "ユーザー: $($entry.UserDisplayName) ($($entry.UserPrincipalName))`n"
    $operationLogContent += "  ステータス: $($entry.Status)`n"
    $operationLogContent += "  処理時刻: $($entry.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))`n"
    if ($entry.ErrorMessage) {
        $operationLogContent += "  エラー: $($entry.ErrorMessage)`n"
    }
    $operationLogContent += "`n"
}

# 操作ログをファイルに記録
Write-Log $operationLogContent "VERBOSE" -NoConsole

}  # <-- この閉じる中括弧が不足していました

# 処理結果の表示
Write-Host "`n処理結果サマリー:" -ForegroundColor Cyan
Write-Host "  成功: $successCount ユーザー" -ForegroundColor Green
Write-Host "  スキップ: $skippedCount ユーザー" -ForegroundColor Yellow
Write-Host "  失敗: $failureCount ユーザー" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
Write-Host "  合計: $($successCount + $skippedCount + $failureCount) ユーザー" -ForegroundColor White

if ($errorDetails.Count -gt 0) {
    Write-Host "`nエラー詳細:" -ForegroundColor Red
    $errorDetails | ForEach-Object { Write-Host "  * $_" -ForegroundColor Red }
    
    # 詳細エラーログにも記録
    Write-Log "==== 処理中のエラー詳細 ====" "ERROR" -NoConsole
    $errorDetails | ForEach-Object {
        Write-Log "  * $_" "ERROR" -NoConsole
    }
}

# 操作の詳細をログに記録
$operationSummary = @"
==== 操作サマリー ====
システム: $($selectedSystem.Name)
パーミッション: $($selectedRole.DisplayName)
アクション: $($actionOptions[$actionChoice])
対象ユーザー数: $($targetUsers.Count)
処理成功: $successCount
処理スキップ: $skippedCount
処理失敗: $failureCount
"@
Write-Log $operationSummary "INFO"

# 実行サマリーをログに記録
Write-ExecutionSummary

# 総合結果ステータスを判定
$resultStatus = if ($failureCount -gt 0) {
    "一部失敗"
} elseif ($skippedCount -gt 0 -and $successCount -eq 0) {
    "全てスキップ"
} elseif ($successCount -gt 0) {
    "成功"
} else {
    "不明"
}

Write-Log "処理総合結果: $resultStatus" $(if ($failureCount -gt 0) { "WARNING" } else { "SUCCESS" })

# スクリプト終了
Write-Host "`n処理が完了しました。" -ForegroundColor Cyan
Write-Host "実況ログは次の場所に保存されました: $logFile" -ForegroundColor Cyan
Write-Host "詳細な情報やエラー内容もログファイルに記録されています。" -ForegroundColor White
Write-Host "Enterキーを押して終了してください..." -ForegroundColor Cyan
Read-Host

# ファイルを右クリックして実行した場合、コンソールウィンドウが即座に閉じないようにする
if ($Host.Name -eq "ConsoleHost") {
    Write-Host "`nこのウィンドウは10秒後に自動的に閉じます..." -ForegroundColor Gray
    Write-Host "閉じたくない場合は、この間に何かキーを押してください..." -ForegroundColor Gray
    
    # キー入力待機または10秒経過
    $startTime = Get-Date
    $timeoutSeconds = 10
    $timeout = New-TimeSpan -Seconds $timeoutSeconds
    
    while (-not [Console]::KeyAvailable -and ((Get-Date) - $startTime) -lt $timeout) {
        Start-Sleep -Milliseconds 200
    }
    
    # キーが押された場合はキーバッファをクリア
    if ([Console]::KeyAvailable) {
        $null = [Console]::ReadKey($true)
    }
}
