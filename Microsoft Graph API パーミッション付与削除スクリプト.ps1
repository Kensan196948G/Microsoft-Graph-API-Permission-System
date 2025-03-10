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

# スクリプト実行開始のログ
$logFile = Join-Path $PSScriptRoot "GraphAPI_Permission_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # コンソールへの出力（色分け）
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage }
    }
    
    # ファイルへの書き込み
    Add-Content -Path $logFile -Value $logMessage
}

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
        $roleAssignments = Get-MgUserMemberOf -UserId (Get-MgUser -UserId "me").Id -ErrorAction Stop
        $globalAdmin = $roleAssignments | Where-Object {$_.DisplayName -eq "Global Administrator"}
        
        if (-not $globalAdmin) {
            Write-Log "グローバル管理者権限がありません。スクリプトを終了します。" "ERROR"
            return $false
        }
        
        Write-Log "グローバル管理者権限が確認されました。" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "管理者権限の確認中にエラーが発生しました: $_" "ERROR"
        return $false
    }
}

function Connect-ToGraph {
    try {
        Write-Log "Microsoft Graph に接続しています..." "INFO"
        Connect-MgGraph -Scopes "Directory.ReadWrite.All", "AppRoleAssignment.ReadWrite.All", "Group.Read.All" -ErrorAction Stop
        Write-Log "Microsoft Graph への接続に成功しました。" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Microsoft Graph への接続に失敗しました: $_" "ERROR"
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
                $foundUsers = Get-MgUser -Filter "startswith(displayName,'$searchQuery') or startswith(userPrincipalName,'$searchQuery') or startswith(onPremisesSamAccountName,'$searchQuery')" -Top 10 -ErrorAction Stop | 
                              Select-Object DisplayName, UserPrincipalName, Id
                
                if ($foundUsers.Count -eq 0) {
                    Write-Log "検索条件に一致するユーザーが見つかりませんでした" "WARNING"
                    return $null
                }
                
                Write-Log "$($foundUsers.Count) 人のユーザーが見つかりました" "INFO"
                
                $userOptions = $foundUsers | ForEach-Object { "$($_.DisplayName) ($($_.UserPrincipalName))" }
                
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
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
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

# 対象のシステムを選択
$systems = @(
    @{ Name = "OneDrive for Business"; AppName = "Office 365 SharePoint Online" },
    @{ Name = "Microsoft Teams"; AppName = "Microsoft Teams Services" },
    @{ Name = "Microsoft EntraID"; AppName = "Microsoft Graph" },
    @{ Name = "Exchange Online"; AppName = "Office 365 Exchange Online" }
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
    "パーミッションを付与",
    "パーミッションを削除",
    "現在のパーミッション割り当てを表示"
)
$actionChoice = Show-Menu -Title "実行するアクション" -Options $actionOptions

if ($actionChoice -eq "Q") {
    Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
    exit 0
}

# 現在のパーミッション割り当てを表示する場合
if ($actionChoice -eq 2) {
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

# 利用可能な API パーミッション（App Role ID）の取得
try {
    $appRoles = $servicePrincipal.AppRoles | Where-Object { $_.IsEnabled -eq $true } | 
                Select-Object DisplayName, Id, Description, Value
    
    if ($appRoles.Count -eq 0) {
        Write-Log "利用可能なAPIパーミッションが見つかりませんでした" "ERROR"
        exit 1
    }
    
    Write-Log "$($appRoles.Count) 個のAPIパーミッションが見つかりました" "INFO"
    
    Write-Host "`n利用可能な API パーミッション:" -ForegroundColor Cyan
    $appRoleOptions = @()
    
    for ($i = 0; $i -lt $appRoles.Count; $i++) {
        $appRoleOptions += "$($appRoles[$i].DisplayName) - $($appRoles[$i].Description)"
        Write-Host "$($i+1). $($appRoles[$i].DisplayName)" -ForegroundColor White
        Write-Host "   ID: $($appRoles[$i].Id)" -ForegroundColor Gray
        Write-Host "   説明: $($appRoles[$i].Description)" -ForegroundColor Gray
        Write-Host "   値: $($appRoles[$i].Value)`n" -ForegroundColor Gray
    }
    
    $roleChoice = Show-Menu -Title "パーミッションを選択" -Options $appRoleOptions
    
    if ($roleChoice -eq "Q") {
        Write-Log "ユーザーによってスクリプトが終了されました" "INFO"
        exit 0
    }
    
    $selectedRole = $appRoles[$roleChoice]
    Write-Log "パーミッション「$($selectedRole.DisplayName)」が選択されました" "INFO"
}
catch {
    Write-Log "APIパーミッションの取得中にエラーが発生しました: $_" "ERROR"
    exit 1
}

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
Write-Host "  パーミッション: $($selectedRole.DisplayName)" -ForegroundColor White
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
$errorDetails = @()

foreach ($user in $targetUsers) {
    try {
        if ($actionChoice -eq 0) {  # パーミッションを付与
            Write-Log "ユーザー「$($user.DisplayName)」にパーミッション「$($selectedRole.DisplayName)」を付与しています..." "INFO"
            
            # 既存の割り当てを確認
            $existingAssignment = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All | 
                                  Where-Object { $_.PrincipalId -eq $user.Id -and $_.AppRoleId -eq $selectedRole.Id }
            
            if ($existingAssignment) {
                Write-Log "ユーザー「$($user.DisplayName)」には既にこのパーミッションが付与されています" "WARNING"
                continue
            }
            
            # パーミッション付与
        
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id `
                -PrincipalId $user.Id `
                -ResourceId $servicePrincipal.Id `
                -AppRoleId $selectedRole.Id -ErrorAction Stop
                
            $successCount++
            Write-Log "ユーザー「$($user.DisplayName)」へのパーミッション付与に成功しました" "SUCCESS"
        }
        elseif ($actionChoice -eq 1) {  # パーミッションを削除
            Write-Log "ユーザー「$($user.DisplayName)」からパーミッション「$($selectedRole.DisplayName)」を削除しています..." "INFO"
            
            # 既存の割り当てを確認
            $existingAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id -All | 
                                 Where-Object { $_.PrincipalId -eq $user.Id -and $_.AppRoleId -eq $selectedRole.Id }
            
            if (-not $existingAssignments -or $existingAssignments.Count -eq 0) {
                Write-Log "ユーザー「$($user.DisplayName)」にはこのパーミッションが付与されていません" "WARNING"
                continue
            }
            
            # 各割り当ての削除
            foreach ($assignment in $existingAssignments) {
                # パーミッション削除（AppRoleAssignmentId で識別）
                Remove-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $servicePrincipal.Id `
                    -AppRoleAssignmentId $assignment.Id -ErrorAction Stop
            }
                
            $successCount++
            Write-Log "ユーザー「$($user.DisplayName)」からのパーミッション削除に成功しました" "SUCCESS"
        }
    }
    catch {
        $failureCount++
        $errorMessage = "ユーザー「$($user.DisplayName)」の処理中にエラーが発生しました: $_"
        $errorDetails += "$($user.UserPrincipalName): $($_.Exception.Message)"
        Write-Log $errorMessage "ERROR"
    }
}

# 処理結果の表示
Write-Host "`n処理結果サマリー:" -ForegroundColor Cyan
Write-Host "  成功: $successCount ユーザー" -ForegroundColor Green
Write-Host "  失敗: $failureCount ユーザー" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })

if ($errorDetails.Count -gt 0) {
    Write-Host "`nエラー詳細:" -ForegroundColor Red
    $errorDetails | ForEach-Object { Write-Host "  * $_" -ForegroundColor Red }
}
Write-Log "スクリプトの実行が完了しました。成功: $successCount, 失敗: $failureCount" $(if ($failureCount -gt 0) { "WARNING" } else { "SUCCESS" })

# スクリプト終了
Write-Host "`n処理が完了しました。ログは次の場所に保存されました: $logFile" -ForegroundColor Cyan
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
