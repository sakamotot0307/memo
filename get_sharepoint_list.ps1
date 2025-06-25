# 認証情報
$tenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxx"

# SharePoint情報
$hostname   = "xxxx.sharepoint.com"
$sitePath   = "/sites/xxxxx"
$listName   = "List"
$outputPath = "SharePointList.csv"

# サイト取得
$siteId = Get-MgSite -Search $sitePath

# リスト取得
$listId = Get-MgSiteList -SiteId $siteId.Id -Filter "DisplayName eq '$listName'"

# リストアイテム取得（fields を展開）
$items = Get-MgSiteListItem -SiteId $siteId.Id -ListId $listId.Id -ExpandProperty "fields" -Top 999

# fields の AdditionalProperties を手動で展開してカスタムオブジェクト化
$data = foreach ($item in $items) {
    $record = @{}
    foreach ($entry in $item.Fields.AdditionalProperties.GetEnumerator()) {
        $record[$entry.Key] = $entry.Value
    }
    [PSCustomObject]$record
}

# import_flg が False のものだけ抽出（文字列化して小文字比較）
$filteredData = $data | Where-Object {
    $_.import_flg.ToString().ToLower() -eq 'false'
}

# CSV 出力（必要なカラムだけ）
$filteredData | Select-Object Title, serialnumber, hostname, import_flg |
    Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Output "CSV 出力完了（import_flg=Falseのみ）: $outputPath"

