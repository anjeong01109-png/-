<#
  호빵 계절별 검색 추이 분석 (NAVER API HUB 데이터랩 검색어 트렌드 + 블로그 검색)
  결과: output\호빵\<날짜_시각>\ 아래 CSV(엑셀용, UTF-8 BOM)와 API 원본 응답(raw\*.json)
  API 키: secrets\naver.json
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$root = Split-Path $PSScriptRoot -Parent
$config = Get-Content (Join-Path $root 'secrets\naver.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$headers = @{ 'X-NCP-APIGW-API-KEY-ID' = $config.clientId; 'X-NCP-APIGW-API-KEY' = $config.clientSecret }
$outDir = Join-Path $root ('output\호빵\' + (Get-Date).ToString('yyyyMMdd_HHmm'))
$rawDir = Join-Path $outDir 'raw'
New-Item -ItemType Directory -Force $rawDir | Out-Null
$utf8 = New-Object Text.UTF8Encoding $false
$today = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')   # 오늘은 집계 중이므로 어제까지

function Invoke-NaverApi([string]$Method, [string]$Path, $Body, [string]$SaveAs) {
    $params = @{ Uri = $config.baseUrl + $Path; Method = $Method; Headers = $headers; UseBasicParsing = $true }
    if ($Body) {
        $params.ContentType = 'application/json; charset=utf-8'
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress))
    }
    $json = [Text.Encoding]::UTF8.GetString((Invoke-WebRequest @params).RawContentStream.ToArray())
    if ($SaveAs) {
        $record = [ordered]@{ request = [ordered]@{ method = $Method; path = $Path; body = $Body }; response = ($json | ConvertFrom-Json) }
        [IO.File]::WriteAllText((Join-Path $rawDir $SaveAs), ($record | ConvertTo-Json -Depth 12), $utf8)
    }
    $json | ConvertFrom-Json
}

function Save-Csv($rows, [string]$name) {
    $csv = $rows | ConvertTo-Csv -NoTypeInformation
    [IO.File]::WriteAllText((Join-Path $outDir $name), ($csv -join "`r`n"), (New-Object Text.UTF8Encoding $true))
}

# 데이터랩은 검색량이 기준 미만인 구간을 응답에서 빼므로, 빠진 구간은 0으로 채움
function Expand-Series($result, [string[]]$periods) {
    $map = @{}; foreach ($d in $result.data) { $map[$d.period] = [double]$d.ratio }
    foreach ($p in $periods) { [pscustomobject]@{ period = $p; ratio = $(if ($map.ContainsKey($p)) { $map[$p] } else { 0.0 }) } }
}

function Get-Periods([datetime]$start, [datetime]$end, [string]$unit) {
    $d = $start
    while ($d -le $end) { $d.ToString('yyyy-MM-dd'); $d = if ($unit -eq 'month') { $d.AddMonths(1) } else { $d.AddDays(7) } }
}

Write-Host '1) 호빵 월간 추이 (2016-01 ~ 현재)'
$r = Invoke-NaverApi 'POST' '/search-trend/v1/search' @{
    startDate = '2016-01-01'; endDate = $today; timeUnit = 'month'
    keywordGroups = @(@{ groupName = '호빵'; keywords = @('호빵') })
} 'trend_monthly_2016_2026.json'
$monthly = Expand-Series $r.results[0] (Get-Periods ([datetime]'2016-01-01') ([datetime]$today) 'month')
Save-Csv ($monthly | ForEach-Object { [pscustomobject]@{ month = $_.period.Substring(0,7); ratio = $_.ratio } }) 'monthly_hoppang_2016_2026.csv'

Write-Host '2) 호빵 주간 추이 (최근 3시즌)'
# 데이터랩 주간 구간은 시작일 기준 월요일 단위로 나뉨 → 응답의 period를 그대로 사용
$r = Invoke-NaverApi 'POST' '/search-trend/v1/search' @{
    startDate = '2023-07-03'; endDate = $today; timeUnit = 'week'
    keywordGroups = @(@{ groupName = '호빵'; keywords = @('호빵') })
} 'trend_weekly_2023_2026.json'
Save-Csv ($r.results[0].data | ForEach-Object { [pscustomobject]@{ week_start = $_.period; ratio = [double]$_.ratio } }) 'weekly_hoppang_2023_2026.csv'

Write-Host '3) 연관 키워드 비교 (최근 3년 월간, 기준 키워드 호빵으로 보정)'
$candidates = @(
    @{ name = '삼립호빵';         type = '브랜드·판매처'; kw = @('삼립호빵','삼립 호빵') }
    @{ name = '편의점 호빵';      type = '브랜드·판매처'; kw = @('편의점 호빵','편의점호빵') }
    @{ name = 'GS25 호빵';        type = '브랜드·판매처'; kw = @('GS25 호빵','gs25호빵','지에스 호빵') }
    @{ name = '포켓몬 호빵';      type = '브랜드·판매처'; kw = @('포켓몬 호빵','포켓몬호빵') }
    @{ name = '호빵 신제품';      type = '브랜드·판매처'; kw = @('호빵 신제품','호빵신제품') }
    @{ name = '단팥호빵';         type = '맛';           kw = @('단팥호빵','단팥 호빵','팥호빵') }
    @{ name = '야채호빵';         type = '맛';           kw = @('야채호빵','야채 호빵') }
    @{ name = '피자호빵';         type = '맛';           kw = @('피자호빵','피자 호빵') }
    @{ name = '김치호빵';         type = '맛';           kw = @('김치호빵','김치 호빵','김치만두호빵') }
    @{ name = '고구마호빵';       type = '맛';           kw = @('고구마호빵','고구마 호빵') }
    @{ name = '카레호빵';         type = '맛';           kw = @('카레호빵','카레 호빵') }
    @{ name = '호빵 찜기';        type = '조리법';       kw = @('호빵 찜기','호빵찜기','호빵 기계','호빵기계') }
    @{ name = '호빵 전자레인지';  type = '조리법';       kw = @('호빵 전자레인지','호빵 전자렌지','호빵전자레인지') }
    @{ name = '호빵 에어프라이어'; type = '조리법';      kw = @('호빵 에어프라이어','호빵에어프라이어') }
    @{ name = '호빵 데우기';      type = '조리법';       kw = @('호빵 데우기','호빵 데우는법','호빵 찌는법','호빵 맛있게 먹는법') }
    @{ name = '호빵 칼로리';      type = '정보';         kw = @('호빵 칼로리','호빵칼로리') }
    @{ name = '호빵 가격';        type = '정보';         kw = @('호빵 가격','호빵가격') }
    @{ name = '붕어빵';           type = '경쟁 겨울간식'; kw = @('붕어빵') }
    @{ name = '호떡';             type = '경쟁 겨울간식'; kw = @('호떡') }
    @{ name = '찐빵';             type = '경쟁 겨울간식'; kw = @('찐빵') }
)
$kStart = [datetime]'2023-10-01'
$kPeriods = Get-Periods $kStart ([datetime]$today) 'month'
$anchor = @{ groupName = '호빵'; keywords = @('호빵') }
$kwRows = @(); $call = 0
for ($i = 0; $i -lt $candidates.Count; $i += 4) {
    $call++
    $chunk = $candidates[$i..([math]::Min($i + 3, $candidates.Count - 1))]
    $groups = @($anchor) + @($chunk | ForEach-Object { @{ groupName = $_.name; keywords = $_.kw } })
    $r = Invoke-NaverApi 'POST' '/search-trend/v1/search' @{
        startDate = $kStart.ToString('yyyy-MM-dd'); endDate = $today; timeUnit = 'month'; keywordGroups = $groups
    } ("trend_keywords_call{0}.json" -f $call)
    # 호출마다 '그 호출 안의 최댓값=100'으로 정규화되므로, 매 호출에 넣은 호빵의 최댓값을 100으로 맞춰 다시 환산
    $anchorSeries = Expand-Series ($r.results | Where-Object title -eq '호빵') $kPeriods
    $factor = 100 / ($anchorSeries | Measure-Object ratio -Maximum).Maximum
    foreach ($res in $r.results) {
        if ($res.title -eq '호빵' -and $call -gt 1) { continue }
        $meta = $candidates | Where-Object { $_.name -eq $res.title }
        foreach ($p in (Expand-Series $res $kPeriods)) {
            $kwRows += [pscustomobject]@{
                keyword = $res.title; type = $(if ($meta) { $meta.type } else { '기준' }); month = $p.period.Substring(0,7)
                ratio_raw = $p.ratio; api_call = $call; ratio_vs_hoppang_peak = [math]::Round($p.ratio * $factor, 3)
            }
        }
    }
    Write-Host "  - call $call : $(($chunk | ForEach-Object name) -join ', ')"
}
Save-Csv $kwRows 'keywords_monthly_2023_2026.csv'

Write-Host '4) 키워드별 블로그 글 수 (콘텐츠 공급량)'
$summary = foreach ($k in (@(@{ name = '호빵'; type = '기준'; kw = @('호빵') }) + $candidates)) {
    $rows = $kwRows | Where-Object keyword -eq $k.name
    $total = ($rows | Measure-Object ratio_vs_hoppang_peak -Sum).Sum
    $peak = $rows | Sort-Object ratio_vs_hoppang_peak -Descending | Select-Object -First 1
    $q = [uri]::EscapeDataString($k.kw[0])
    $blog = (Invoke-NaverApi 'GET' "/search/v1/blog?query=$q&display=1" $null $null).total
    $season = ($rows | Where-Object { $_.month.Substring(5) -in '11','12','01','02' } | Measure-Object ratio_vs_hoppang_peak -Sum).Sum
    [pscustomobject]@{
        keyword = $k.name; type = $k.type; variants = ($k.kw -join ' | ')
        search_index_3y_sum = [math]::Round($total, 1)
        winter_share_pct = $(if ($total) { [math]::Round($season / $total * 100, 1) } else { 0 })
        peak_month = $peak.month; blog_posts = [int64]$blog
    }
}
Save-Csv ($summary | Sort-Object search_index_3y_sum -Descending) 'keywords_summary.csv'

Write-Host "`n완료: $outDir"
