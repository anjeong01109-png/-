<#
  행사·전시·브랜드 체험 마케팅 대행사 비교 (추천 후보 정리용)

  용도: 후보 대행사별로 ① 검색 지수(찾는 사람) ② 최근 12개월 뉴스 수 ③ 그중 행사·전시·팝업 관련 기사 수
        ④ 블로그 글 수(전체)를 모아 한 표로 비교합니다.
  사용 예:
    .\analysis\event-agency-analysis.ps1
  결과: output\행사대행사-추천\<yyyymmdd_hhmm>\
        agencies_summary.csv   대행사별 비교표
        trend_monthly.csv      월별 검색 지수 (기준 키워드로 환산)
        news_items.csv         대행사 이름이 들어간 최근 12개월 뉴스 목록
        raw\*.json             요청 내용과 응답 원문
  API 키: secrets\naver.json
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$root = Split-Path $PSScriptRoot -Parent
$config = Get-Content (Join-Path $root 'secrets\naver.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$headers = @{
    'X-NCP-APIGW-API-KEY-ID' = $config.clientId
    'X-NCP-APIGW-API-KEY'    = $config.clientSecret
}

# 후보: group = 구분, query = 검색어, match = 제목·요약에 이 중 하나가 있어야 해당 회사 글로 인정
$agencies = @(
    @{ slug = 'cheil';      name = '제일기획';           group = '대형 종합광고';     query = '제일기획';           match = @('제일기획') }
    @{ slug = 'innocean';   name = '이노션';             group = '대형 종합광고';     query = '이노션';             match = @('이노션') }
    @{ slug = 'hsad';       name = 'HS애드';             group = '대형 종합광고';     query = 'HS애드';             match = @('HS애드','HSAD') }
    @{ slug = 'daehong';    name = '대홍기획';           group = '대형 종합광고';     query = '대홍기획';           match = @('대홍기획') }
    @{ slug = 'estylelmc';  name = '이스타일엘엠씨';     group = '행사·전시 전문';    query = '이스타일엘엠씨';     match = @('이스타일엘엠씨','이스타일 엘엠씨','E.STYLE LMC') }
    @{ slug = 'eidetic';    name = '아이데틱';           group = '행사·전시 전문';    query = '아이데틱 마케팅';    match = @('아이데틱','EIDETIC') }
    @{ slug = 'tradeworld'; name = '트레이드월드';       group = '행사·전시 전문';    query = '트레이드월드';       match = @('트레이드월드') }
    @{ slug = 'sigong';     name = '시공테크';           group = '행사·전시 전문';    query = '시공테크';           match = @('시공테크') }
    @{ slug = 'unione';     name = '유니원커뮤니케이션즈'; group = '행사·전시 전문';  query = '유니원커뮤니케이션즈'; match = @('유니원커뮤니케이션즈','유니원') }
    @{ slug = 'fuldalab';   name = '풀다랩';             group = '팝업·브랜드 체험';  query = '풀다랩';             match = @('풀다랩') }
    @{ slug = 'xpcraft';    name = '엑스피크래프트';     group = '팝업·브랜드 체험';  query = '엑스피크래프트';     match = @('엑스피크래프트') }
    @{ slug = 'madecorp';   name = '메이드코프';         group = '팝업·브랜드 체험';  query = '메이드코프';         match = @('메이드코프') }
    @{ slug = 'sweetspot';  name = '스위트스팟';         group = '팝업·브랜드 체험';  query = '스위트스팟';         match = @('스위트스팟') }
    @{ slug = 'dminusone';  name = '디마이너스원';       group = '팝업·브랜드 체험';  query = '디마이너스원';       match = @('디마이너스원') }
)
$eventWords = '팝업|전시|행사|부스|박람회|체험|페스티벌|축제|이벤트|쇼케이스|컨벤션'

$now = Get-Date
$stamp = $now.ToString('yyyyMMdd_HHmm')
$outDir = Join-Path $root "output\행사대행사-추천\$stamp"
$rawDir = Join-Path $outDir 'raw'
New-Item -ItemType Directory -Force $rawDir | Out-Null
$utf8 = New-Object Text.UTF8Encoding $false
$utf8Bom = New-Object Text.UTF8Encoding $true

function Invoke-NaverApi([string]$Method, [string]$Path, $Body) {
    $params = @{ Uri = $config.baseUrl + $Path; Method = $Method; Headers = $headers; UseBasicParsing = $true }
    if ($Body) {
        $params.ContentType = 'application/json; charset=utf-8'
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress))
    }
    $resp = Invoke-WebRequest @params
    [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
}

# 호출 + 원문 저장. 실패하면 상태 번호를 남기고 멈추지 않음
function Get-Saved([string]$File, [string]$Method, [string]$Path, $Body) {
    $rec = [ordered]@{ request = @{ method = $Method; path = $Path; body = $Body }; status = 200; response = $null }
    try {
        $text = Invoke-NaverApi $Method $Path $Body
        $rec.response = $text | ConvertFrom-Json
    } catch {
        $rec.status = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
        $rec.error = $_.Exception.Message
    }
    [IO.File]::WriteAllText((Join-Path $rawDir $File), ($rec | ConvertTo-Json -Depth 12), $utf8)
    [pscustomobject]$rec
}

function Clear-Html([string]$s) {
    if (-not $s) { return '' }
    [Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', '')).Trim()
}

function Test-Match([string]$text, [string[]]$needles) {
    $t = $text -replace '\s', ''
    foreach ($n in $needles) { if ($t -like "*$($n -replace '\s','')*") { return $true } }
    $false
}

# 1) 검색어 트렌드: 최근 12개 완료월 (이번 달은 집계 중이라 제외). 매 호출에 기준 키워드를 넣어 환산
$trendStart = $now.AddMonths(-12).ToString('yyyy-MM-01')
$trendEnd = (Get-Date -Day 1).AddDays(-1).ToString('yyyy-MM-dd')
$baseline = $agencies | Where-Object slug -eq 'innocean'
$others = @($agencies | Where-Object slug -ne 'innocean')
$months = @()
$trendRaw = @{}   # slug -> @{ month -> ratio(호출 내) }
$baseRaw = @()    # 호출별 기준 키워드 월 값
for ($i = 0; $i -lt $others.Count; $i += 4) {
    $chunk = $others[$i..([math]::Min($i + 3, $others.Count - 1))]
    $groups = @(@{ groupName = $baseline.slug; keywords = @($baseline.match) })
    foreach ($a in $chunk) { $groups += @{ groupName = $a.slug; keywords = @($a.match | Select-Object -First 5) } }
    $body = @{ startDate = $trendStart; endDate = $trendEnd; timeUnit = 'month'; keywordGroups = $groups }
    $n = [int]($i / 4) + 1
    $r = Get-Saved "trend_call$n.json" 'POST' '/search-trend/v1/search' $body
    if ($r.status -ne 200) { Write-Host "  - 트렌드 호출 $n 실패: $($r.status)"; continue }
    if (-not $months) {
        $d = [datetime]$trendStart
        while ($d -le [datetime]$trendEnd) { $months += $d.ToString('yyyy-MM-01'); $d = $d.AddMonths(1) }
    }
    $callBase = $null
    foreach ($res in $r.response.results) {
        $map = @{}; foreach ($m in $months) { $map[$m] = 0.0 }   # 빠진 달 = 기준 미만 → 0
        foreach ($p in $res.data) { $map[$p.period] = [double]$p.ratio }
        if ($res.title -eq $baseline.slug) { $callBase = $map } else { $trendRaw[$res.title] = @{ map = $map; call = $n } }
    }
    $baseRaw += , @{ call = $n; map = $callBase }
    Write-Host "  - 트렌드 호출 $n OK"
}
# 첫 호출의 기준 키워드 합계에 맞춰 각 호출 값을 환산
$refSum = ($baseRaw[0].map.Values | Measure-Object -Sum).Sum
$scaled = @{}
$scaled[$baseline.slug] = $baseRaw[0].map
foreach ($slug in $trendRaw.Keys) {
    $call = $trendRaw[$slug].call
    $bs = ($baseRaw | Where-Object { $_.call -eq $call }).map
    $factor = $refSum / (($bs.Values | Measure-Object -Sum).Sum)
    $m2 = @{}; foreach ($m in $months) { $m2[$m] = $trendRaw[$slug].map[$m] * $factor }
    $scaled[$slug] = $m2
}
# 전체 최댓값 = 100 으로 다시 맞춤
$max = ($scaled.Values | ForEach-Object { $_.Values } | Measure-Object -Maximum).Maximum
foreach ($slug in @($scaled.Keys)) { foreach ($m in $months) { $scaled[$slug][$m] = [math]::Round($scaled[$slug][$m] / $max * 100, 2) } }

# 2) 뉴스(최신순 100건 → 최근 12개월·이름 포함만) + 블로그 글 수
$since = $now.AddYears(-1)
$summary = @(); $newsRows = @()
foreach ($a in $agencies) {
    $q = [uri]::EscapeDataString($a.query)
    $news = Get-Saved "news_$($a.slug).json" 'GET' "/search/v1/news?query=$q&display=100&sort=date&format=json" $null
    $blog = Get-Saved "blog_$($a.slug).json" 'GET' "/search/v1/blog?query=$q&display=10&sort=sim&format=json" $null
    $items = @()
    if ($news.status -eq 200) {
        foreach ($it in $news.response.items) {
            $title = Clear-Html $it.title; $desc = Clear-Html $it.description
            if (-not (Test-Match "$title $desc" $a.match)) { continue }
            $date = [datetime]$it.pubDate
            if ($date -lt $since) { continue }
            $isEvent = ("$title $desc" -match $eventWords)
            $items += [pscustomobject]@{ agency = $a.name; date = $date.ToString('yyyy-MM-dd'); event_related = $isEvent; title = $title; link = $(if ($it.originallink) { $it.originallink } else { $it.link }) }
        }
    }
    $newsRows += $items
    $oldest = if ($news.status -eq 200 -and $news.response.items.Count) { ([datetime]($news.response.items[-1].pubDate)).ToString('yyyy-MM-dd') } else { '' }
    $capped = ($news.status -eq 200 -and $news.response.items.Count -eq 100 -and [datetime]($news.response.items[-1].pubDate) -ge $since)
    $t = $scaled[$a.slug]
    $summary += [pscustomobject]@{
        group                 = $a.group
        agency                = $a.name
        search_index_avg_12m  = $(if ($t) { [math]::Round((($t.Values | Measure-Object -Average).Average), 2) } else { '' })
        news_12m              = $items.Count
        news_12m_capped       = $capped
        news_12m_event        = @($items | Where-Object event_related).Count
        news_total            = $(if ($news.status -eq 200) { $news.response.total } else { "오류 $($news.status)" })
        blog_total            = $(if ($blog.status -eq 200) { $blog.response.total } else { "오류 $($blog.status)" })
        oldest_news_in_fetch  = $oldest
    }
    Write-Host "  - $($a.name): 뉴스 $($items.Count)건(12개월), 행사 관련 $(@($items | Where-Object event_related).Count)건"
}

# 3) 저장 (CSV는 엑셀용 BOM UTF-8)
function Save-Csv($rows, [string]$file) {
    $csv = ($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    [IO.File]::WriteAllText((Join-Path $outDir $file), $csv, $utf8Bom)
}
Save-Csv $summary 'agencies_summary.csv'
Save-Csv $newsRows 'news_items.csv'
$trendRows = foreach ($m in $months) {
    $row = [ordered]@{ month = $m.Substring(0, 7) }
    foreach ($a in $agencies) { $row[$a.name] = $(if ($scaled[$a.slug]) { $scaled[$a.slug][$m] } else { '' }) }
    [pscustomobject]$row
}
Save-Csv $trendRows 'trend_monthly.csv'
$info = [ordered]@{
    collectedAt   = $now.ToString('s')
    trendPeriod   = "$trendStart ~ $trendEnd (월간, 이번 달 제외)"
    trendBaseline = $baseline.name
    newsWindow    = "$($since.ToString('yyyy-MM-dd')) ~ $($now.ToString('yyyy-MM-dd')) (최신순 최대 100건 안에서)"
    eventWords    = $eventWords
}
[IO.File]::WriteAllText((Join-Path $outDir 'collection_info.json'), ($info | ConvertTo-Json), $utf8)

Write-Host "`n완료: $outDir" -ForegroundColor Green
