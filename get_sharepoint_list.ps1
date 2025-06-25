# 認証情報
$tenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxx"

# SharePoint情報
$hostname   = "xxxx.sharepoint.com"
$sitePath   = "/sites/xxxxx"
$listName   = "List"
$outputPath = "SharePointList.csv"

$Client_Secret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $Client_Secret
Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $ClientSecretCredential

# サイトとリストの取得
$siteId = Get-MgSite -Search $sitePath
$listId = Get-MgSiteList -SiteId $siteId.Id -Filter "DisplayName eq '$listName'"

# リストアイテム取得（fields を展開）
$items = Get-MgSiteListItem -SiteId $siteId.Id -ListId $listId.Id -ExpandProperty "fields" -All

# fields の AdditionalProperties を手動で展開してカスタムオブジェクト化
$data = foreach ($item in $items) {
    $record = @{}

    # fields の各プロパティを追加
    foreach ($entry in $item.Fields.AdditionalProperties.GetEnumerator()) {
        $record[$entry.Key] = $entry.Value
    }

    # アイテムIDとETagを追加
    $record["item_id"] = $item.Id
    $record["etag"]    = $item.ETag

    [PSCustomObject]$record
}

# import_flg が False のものだけ抽出（文字列化して小文字比較）
$filteredData = $data | Where-Object {
    $_.import_flg.ToString().ToLower() -eq 'false'
}

$filteredData | Select-Object `
    item_id, etag, `
    Title, serialnumber, hostname, import_flg |
    Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Output "CSV 出力完了（import_flg=Falseのみ）: $outputPath"
