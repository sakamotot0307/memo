# 認証情報
$tenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxx"

# SharePoint情報
$hostname   = "xxxx.sharepoint.com"
$sitePath   = "/sites/xxxxx"
$listName   = "List"
$inputPath = "SharePointList.csv"

$Client_Secret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $Client_Secret
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential

# サイトとリストの取得
$siteId = Get-MgSite -Search $sitePath
$listId = Get-MgSiteList -SiteId $siteId.Id -Filter "DisplayName eq '$listName'"

# CSV読み込み
$csv = Import-Csv -Path $inputPath

# CSVの各行を更新
foreach ($row in $csv) {
    $itemId = $row.item_id

    if (![string]::IsNullOrEmpty($itemId)) {
        Write-Output "Updating item_id=$itemId..."

        $params = @{
            fields = @{
                import_flg = $true
            }
        }

        try {
            Update-MgSiteListItem -SiteId $siteId.Id -ListId $listId.Id -ListItemId $itemId -BodyParameter $params
            Write-Output "item_id=$itemId 更新完了"
        } catch {
            Write-Warning "item_id=$itemId の更新に失敗: $_"
        }
    } else {
        Write-Warning "item_id が空です。スキップします。"
    }
}

Write-Output "すべての更新が完了しました。"
