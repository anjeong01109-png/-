<#
  NAVER API HUB 주제 리서치 도구 (콘텐츠 기획/초안용)

  사용 예:
    .\naver-topic.ps1 -Topic "홈카페"
    .\naver-topic.ps1 -Topic "홈카페" -Compare "캠핑","러닝" -Months 24 -Count 20

  결과: output\<주제>\<날짜_시각>\report.md (사람이 읽는 리포트), raw.json (원본 데이터)
  API 키: secrets\naver.json
#>
param(
    [string]$Topic,
    [string[]]$Compare = @(),   # 트렌드 비교용 키워드 (최대 4개)
    [int]$Months = 12,          # 검색어 트렌드 조회 기간(개월)
    [int]$Count = 10            # 채널별 가져올 결과 수 (1~100)
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not $Topic) { $Topic = Read-Host '리서치할 주제를 입력하세요' }
if (-not $Topic) { throw '주제가 비어 있습니다.' }

$root = Split-Path $PSScriptRoot -Parent   # 프로젝트 폴더 (tools의 한 단계 위)
$config = Get-Content (Join-Path $root 'secrets\naver.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$headers = @{
    'X-NCP-APIGW-API-KEY-ID' = $config.clientId
    'X-NCP-APIGW-API-KEY'    = $config.clientSecret
}

# PowerShell 5.1의 Invoke-RestMethod는 charset 없는 응답을 깨뜨리므로 UTF-8로 직접 디코딩
function Invoke-NaverApi([string]$Method, [string]$Path, $Body) {
    $params = @{ Uri = $config.baseUrl + $Path; Method = $Method; Headers = $headers; UseBasicParsing = $true }
    if ($Body) {
        $params.ContentType = 'application/json; charset=utf-8'
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress))
    }
    $resp = Invoke-WebRequest @params
    [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray()) | ConvertFrom-Json
}

function Get-ErrorText($err) {
    $code = $null
    if ($err.Exception.Response) { $code = [int]$err.Exception.Response.StatusCode }
    if ($code -eq 401) { return '401 권한 없음 - NCP 콘솔 > NAVER API HUB > Application에서 이 API 사용을 추가하세요.' }
    return "$code $($err.Exception.Message)"
}

function Clear-Html([string]$s) {
    if (-not $s) { return '' }
    $s = $s -replace '<[^>]+>', ''
    [Net.WebUtility]::HtmlDecode($s).Trim()
}

function Search-Naver([string]$Type, [string]$Sort = 'sim') {
    $q = [uri]::EscapeDataString($Topic)
    try {
        $r = Invoke-NaverApi 'GET' "/search/v1/$($Type)?query=$q&display=$Count&sort=$Sort&format=json"
        [pscustomobject]@{ ok = $true; total = $r.total; items = @($r.items); error = $null }
    } catch {
        [pscustomobject]@{ ok = $false; total = 0; items = @(); error = (Get-ErrorText $_) }
    }
}

Write-Host "`n[$Topic] 데이터 수집 중..." -ForegroundColor Cyan

# 1) 검색 채널별 데이터
$channels = [ordered]@{
    blog        = @{ name = '블로그 (정확도순)'; sort = 'sim' }
    blogRecent  = @{ name = '블로그 (최신순)';   sort = 'date'; type = 'blog' }
    news        = @{ name = '뉴스 (최신순)';     sort = 'date' }
    cafearticle = @{ name = '카페글';            sort = 'sim' }
    kin         = @{ name = '지식iN (독자 질문)'; sort = 'sim' }
    webkr       = @{ name = '웹문서';            sort = 'sim' }
    encyc       = @{ name = '백과사전';          sort = 'sim' }
}
$search = [ordered]@{}
foreach ($key in $channels.Keys) {
    $ch = $channels[$key]
    $type = if ($ch.type) { $ch.type } else { $key }
    $search[$key] = Search-Naver $type $ch.sort
    $status = if ($search[$key].ok) { "OK ($($search[$key].total)건)" } else { "건너뜀: $($search[$key].error)" }
    Write-Host ("  - {0}: {1}" -f $ch.name, $status)
}

# 최신순 블로그는 주제와 느슨하게만 일치하는 글이 섞이므로 제목/요약에 주제어(띄어쓰기 단위 전부)가 있는 글만 남김
$needles = $Topic -split '\s+' | Where-Object { $_ }
$search.blogRecent.items = @($search.blogRecent.items | Where-Object {
    $t = ((Clear-Html $_.title) + (Clear-Html $_.description)) -replace '\s', ''
    @($needles | Where-Object { $t -notlike "*$_*" }).Count -eq 0
})

# 2) 데이터랩 검색어 트렌드 (주제 + 비교 키워드, 같은 호출 안에서만 비율 비교 가능)
$end = Get-Date
$start = $end.AddMonths(-$Months)
$groups = @(@{ groupName = $Topic; keywords = @($Topic) })
foreach ($c in ($Compare | Select-Object -First 4)) { $groups += @{ groupName = $c; keywords = @($c) } }
$trend = $null; $trendError = $null
try {
    $trend = Invoke-NaverApi 'POST' '/search-trend/v1/search' @{
        startDate = $start.ToString('yyyy-MM-01'); endDate = $end.ToString('yyyy-MM-dd')
        timeUnit = 'month'; keywordGroups = $groups
    }
    Write-Host '  - 검색어 트렌드: OK'
} catch {
    $trendError = Get-ErrorText $_
    Write-Host "  - 검색어 트렌드: 실패 ($trendError)"
}

# 3) 블로그/카페/지식iN 제목·요약에서 자주 나오는 표현 (형태소 분석 없는 단순 빈도)
$stop = @('있는','있습니다','합니다','하는','그리고','하지만','정말','너무','이번','오늘','많이','같은','위한','대한','통해','에서','으로','입니다','있어요','하고','해서','안녕하세요','여러분','그래서','이렇게','어떻게','진짜','먼저','바로','우리','지금')
$text = foreach ($k in 'blog','blogRecent','cafearticle','kin') {
    foreach ($it in $search[$k].items) { (Clear-Html $it.title) + ' ' + (Clear-Html $it.description) }
}
$words = ($text -join ' ') -split '[^\p{L}\p{Nd}]+' |
    ForEach-Object { if ($_.Length -ge 3) { $_ -replace '(을|를|은|는|이|가|의|에|도|과|와|로|으로|에서)$', '' } else { $_ } } |
    Where-Object { $_.Length -ge 2 -and $_ -ne $Topic -and $stop -notcontains $_ -and $_ -notmatch '^\d+$' } |
    Group-Object | Where-Object Count -ge 2 | Sort-Object Count -Descending | Select-Object -First 30

# 4) 리포트 작성
$md = New-Object Collections.Generic.List[string]
$md.Add("# 주제 리서치: $Topic")
$md.Add('')
$md.Add("- 수집 시각: $($end.ToString('yyyy-MM-dd HH:mm'))")
$md.Add("- 검색량 규모(총 결과 수): " + (($channels.Keys | Where-Object { $search[$_].ok -and $_ -ne 'blogRecent' } |
    ForEach-Object { "$($channels[$_].name -replace ' \(.*\)','') $('{0:N0}' -f [int64]$search[$_].total)" }) -join ' · '))
$md.Add('')

$md.Add("## 검색어 트렌드 (최근 $($Months)개월, 월간)")
$md.Add('')
if ($trend) {
    $md.Add('> 수치는 조회 기간 중 최대 검색량을 100으로 둔 상대값입니다.')
    $md.Add('')
    foreach ($res in $trend.results) {
        $data = @($res.data)
        $md.Add("**$($res.title)**")
        $md.Add('')
        $md.Add('| 월 | 지수 | |')
        $md.Add('|---|---:|---|')
        foreach ($d in $data) {
            $bar = '█' * [math]::Round([double]$d.ratio / 5)
            $md.Add("| $($d.period.Substring(0,7)) | $('{0:N1}' -f [double]$d.ratio) | $bar |")
        }
        if ($data.Count -ge 6) {
            $recent = ($data | Select-Object -Last 3 | Measure-Object ratio -Average).Average
            $prev = ($data | Select-Object -Last 6 | Select-Object -First 3 | Measure-Object ratio -Average).Average
            $peak = $data | Sort-Object { [double]$_.ratio } -Descending | Select-Object -First 1
            $dir = if ($prev -eq 0) { '-' } else { '{0:+0;-0}%' -f (($recent - $prev) / $prev * 100) }
            $md.Add('')
            $md.Add("- 최근 3개월 평균 vs 직전 3개월: **$dir** · 최고점: $($peak.period.Substring(0,7))")
        }
        $md.Add('')
    }
} else {
    $md.Add("트렌드 조회 실패: $trendError")
    $md.Add('')
}

$md.Add('## 자주 함께 나오는 표현 (블로그·카페·지식iN 기준)')
$md.Add('')
$md.Add(($words | ForEach-Object { "$($_.Name)($($_.Count))" }) -join ', ')
$md.Add('')

if ($search.kin.ok -and $search.kin.items.Count) {
    $md.Add('## 독자들이 실제로 묻는 질문 (지식iN)')
    $md.Add('')
    foreach ($it in $search.kin.items) { $md.Add("- [$(Clear-Html $it.title)]($($it.link))") }
    $md.Add('')
}

foreach ($key in $channels.Keys) {
    if ($key -eq 'kin') { continue }
    $r = $search[$key]
    $md.Add("## $($channels[$key].name)")
    $md.Add('')
    if (-not $r.ok) { $md.Add("_$($r.error)_"); $md.Add(''); continue }
    if (-not $r.items.Count) { $md.Add('_결과 없음_'); $md.Add(''); continue }
    foreach ($it in $r.items) {
        $meta = @()
        if ($it.postdate) { $meta += $it.postdate.Insert(6,'-').Insert(4,'-') }
        if ($it.pubDate) { $meta += ([datetime]$it.pubDate).ToString('yyyy-MM-dd') }
        if ($it.bloggername) { $meta += $it.bloggername }
        if ($it.cafename) { $meta += $it.cafename }
        $suffix = if ($meta) { " _($($meta -join ' · '))_" } else { '' }
        $link = if ($it.originallink) { $it.originallink } else { $it.link }
        $md.Add("- **[$(Clear-Html $it.title)]($link)**$suffix")
        $desc = Clear-Html $it.description
        if ($desc) { $md.Add("  - $desc") }
    }
    $md.Add('')
}

# 5) 저장
$safe = ($Topic -replace '[\\/:*?"<>|\s]+', '_')
$outDir = Join-Path $root ("output\{0}\{1}" -f $safe, $end.ToString('yyyyMMdd_HHmm'))
New-Item -ItemType Directory -Force $outDir | Out-Null
$utf8 = New-Object Text.UTF8Encoding $false
[IO.File]::WriteAllText((Join-Path $outDir 'report.md'), ($md -join "`r`n"), $utf8)
$raw = [ordered]@{ topic = $Topic; collectedAt = $end.ToString('s'); search = $search; trend = $trend }
[IO.File]::WriteAllText((Join-Path $outDir 'raw.json'), ($raw | ConvertTo-Json -Depth 10), $utf8)

Write-Host "`n완료: $outDir" -ForegroundColor Green
Write-Host '  - report.md : 리서치 리포트'
Write-Host '  - raw.json  : 원본 데이터 (초안 작성 시 Claude에게 그대로 전달 가능)'
