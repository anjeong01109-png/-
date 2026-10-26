<#
  젝시믹스 vs 안다르 경쟁사 분석 (NAVER API HUB 검색어 트렌드 + 블로그·뉴스 검색)

  사용 예:  .\xexymix-andar-analysis.ps1
  결과: output\젝시믹스_vs_안다르\<날짜_시각>\ 아래 CSV(엑셀용, UTF-8 BOM)와 API 원본 응답(raw\*.json)
  API 키: secrets\naver.json
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$root = Split-Path $PSScriptRoot -Parent
$config = Get-Content (Join-Path $root 'secrets\naver.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$headers = @{ 'X-NCP-APIGW-API-KEY-ID' = $config.clientId; 'X-NCP-APIGW-API-KEY' = $config.clientSecret }
$now = Get-Date
$outDir = Join-Path $root ('output\젝시믹스_vs_안다르\' + $now.ToString('yyyyMMdd_HHmm'))
$rawDir = Join-Path $outDir 'raw'
New-Item -ItemType Directory -Force $rawDir | Out-Null
$utf8 = New-Object Text.UTF8Encoding $false

# 집계가 끝난 구간만 사용: 월간은 지난달까지, 주간은 지난주 일요일까지
$lastMonthEnd = $now.Date.AddDays(1 - $now.Day).AddDays(-1)
$y1Start = $lastMonthEnd.AddDays(1).AddMonths(-12)
$thisMonday = $now.Date.AddDays(-(([int]$now.DayOfWeek + 6) % 7))
$lastSunday = $thisMonday.AddDays(-1)
$weekStart = $thisMonday.AddDays(-7 * 52)
function Fmt([datetime]$d) { $d.ToString('yyyy-MM-dd') }

function Invoke-NaverApi([string]$Method, [string]$Path, $Body, [string]$SaveAs) {
    $params = @{ Uri = $config.baseUrl + $Path; Method = $Method; Headers = $headers; UseBasicParsing = $true }
    if ($Body) {
        $params.ContentType = 'application/json; charset=utf-8'
        $params.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10 -Compress))
    }
    $json = [Text.Encoding]::UTF8.GetString((Invoke-WebRequest @params).RawContentStream.ToArray())
    if ($SaveAs) {
        $record = [ordered]@{ request = [ordered]@{ method = $Method; path = $Path; body = $Body; collectedAt = $now.ToString('s') }; response = ($json | ConvertFrom-Json) }
        [IO.File]::WriteAllText((Join-Path $rawDir $SaveAs), ($record | ConvertTo-Json -Depth 12), $utf8)
    }
    $json | ConvertFrom-Json
}

function Save-Csv($rows, [string]$name) {
    $csv = $rows | ConvertTo-Csv -NoTypeInformation
    [IO.File]::WriteAllText((Join-Path $outDir $name), ($csv -join "`r`n"), (New-Object Text.UTF8Encoding $true))
}

function Clear-Html([string]$s) {
    if (-not $s) { return '' }
    [Net.WebUtility]::HtmlDecode(($s -replace '<[^>]+>', '')).Trim()
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

function Get-Trend($groups, [datetime]$start, [datetime]$end, [string]$unit, [string]$saveAs, $extra) {
    $body = @{ startDate = (Fmt $start); endDate = (Fmt $end); timeUnit = $unit; keywordGroups = $groups }
    if ($extra) { foreach ($k in $extra.Keys) { $body[$k] = $extra[$k] } }
    Invoke-NaverApi 'POST' '/search-trend/v1/search' $body $saveAs
}

$brands = [ordered]@{
    '젝시믹스' = @{ kw = @('젝시믹스', 'xexymix', '젝시 믹스'); match = '젝시믹스|젝시 믹스|xexymix'; file = 'xexymix' }
    '안다르'   = @{ kw = @('안다르', 'andar');                  match = '안다르|andar';             file = 'andar' }
}
$brandGroups = @($brands.Keys | ForEach-Object { @{ groupName = $_; keywords = $brands[$_].kw } })

Write-Host '1) 브랜드 월간 검색 추이 (2018-01 ~ 지난달, 비교용 일반어: 레깅스)'
$mStart = [datetime]'2018-01-01'
$r = Get-Trend ($brandGroups + @(@{ groupName = '레깅스'; keywords = @('레깅스') })) $mStart $lastMonthEnd 'month' 'trend_monthly.json' $null
$mPeriods = @(Get-Periods $mStart $lastMonthEnd 'month')
$series = @{}; foreach ($res in $r.results) { $series[$res.title] = @(Expand-Series $res $mPeriods) }
Save-Csv (0..($mPeriods.Count - 1) | ForEach-Object {
    [pscustomobject]@{ month = $mPeriods[$_].Substring(0, 7); '젝시믹스' = $series['젝시믹스'][$_].ratio; '안다르' = $series['안다르'][$_].ratio; '레깅스' = $series['레깅스'][$_].ratio }
}) 'trend_monthly_2018_now.csv'

Write-Host '2) 브랜드 주간 검색 추이 (최근 52주)'
$r = Get-Trend $brandGroups $weekStart $lastSunday 'week' 'trend_weekly.json' $null
$wPeriods = @(Get-Periods $weekStart $lastSunday 'week')
$series = @{}; foreach ($res in $r.results) { $series[$res.title] = @(Expand-Series $res $wPeriods) }
Save-Csv (0..($wPeriods.Count - 1) | ForEach-Object {
    [pscustomobject]@{ week_start = $wPeriods[$_]; '젝시믹스' = $series['젝시믹스'][$_].ratio; '안다르' = $series['안다르'][$_].ratio }
}) 'trend_weekly_52w.csv'

Write-Host '3) 성별·연령·기기별 검색 점유율 (최근 12개월)'
# 구간마다 별도 호출이므로 구간끼리 크기 비교는 불가. 같은 호출 안의 두 브랜드 비율(점유율)만 비교
$segments = @(
    @{ name = '전체';        extra = $null }
    @{ name = '여성';        extra = @{ gender = 'f' } }
    @{ name = '남성';        extra = @{ gender = 'm' } }
    @{ name = '13-18세';     extra = @{ ages = @('2') } }
    @{ name = '19-29세';     extra = @{ ages = @('3', '4') } }
    @{ name = '30대';        extra = @{ ages = @('5', '6') } }
    @{ name = '40대';        extra = @{ ages = @('7', '8') } }
    @{ name = '50세 이상';   extra = @{ ages = @('9', '10', '11') } }
    @{ name = '모바일';      extra = @{ device = 'mo' } }
    @{ name = 'PC';          extra = @{ device = 'pc' } }
)
$yPeriods = @(Get-Periods $y1Start $lastMonthEnd 'month')
$segRows = foreach ($s in $segments) {
    $r = Get-Trend $brandGroups $y1Start $lastMonthEnd 'month' ("trend_segment_{0}.json" -f [array]::IndexOf($segments, $s)) $s.extra
    $sum = @{}; foreach ($res in $r.results) { $sum[$res.title] = (Expand-Series $res $yPeriods | Measure-Object ratio -Sum).Sum }
    $x = $sum['젝시믹스']; $a = $sum['안다르']
    [pscustomobject]@{ segment = $s.name; xexymix_index_sum = [math]::Round($x, 1); andar_index_sum = [math]::Round($a, 1)
        xexymix_share_pct = $(if ($x + $a) { [math]::Round($x / ($x + $a) * 100, 1) } else { 0 }) }
}
Save-Csv $segRows 'trend_segments_12m.csv'

Write-Host '4) 관심 주제별 검색 지수와 블로그 글 수 (최근 12개월)'
$topics = @(
    @{ name = '레깅스'; suf = @('레깅스') }
    @{ name = '골프';   suf = @('골프', '골프웨어') }
    @{ name = '남성';   suf = @('남자', '남성', '맨즈') }
    @{ name = '수영복'; suf = @('수영복', '스윔') }
    @{ name = '후기';   suf = @('후기', '리뷰') }
    @{ name = '할인';   suf = @('할인', '세일', '쿠폰') }
    @{ name = '사이즈'; suf = @('사이즈') }
    @{ name = '매장';   suf = @('매장', '오프라인') }
)
# 매 호출에 기준 키워드(젝시믹스 브랜드 전체)를 넣고, 그 최댓값을 100으로 맞춰 호출 간 값을 환산
$anchor = @{ groupName = '기준'; keywords = $brands['젝시믹스'].kw }
$subRows = @(); $call = 0
for ($i = 0; $i -lt $topics.Count; $i += 2) {
    $call++
    $groups = @($anchor)
    foreach ($t in $topics[$i..($i + 1)]) {
        foreach ($b in $brands.Keys) {
            $groups += @{ groupName = "$b|$($t.name)"; keywords = @($t.suf | ForEach-Object { "$b $_"; "$b$_" }) }
        }
    }
    $r = Get-Trend $groups $y1Start $lastMonthEnd 'month' ("trend_topics_call{0}.json" -f $call) $null
    $factor = 100 / (Expand-Series ($r.results | Where-Object title -eq '기준') $yPeriods | Measure-Object ratio -Maximum).Maximum
    foreach ($res in ($r.results | Where-Object title -ne '기준')) {
        $b, $t = $res.title -split '\|'
        $suf = ($topics | Where-Object name -eq $t).suf[0]
        $total = (Invoke-NaverApi 'GET' "/search/v1/blog?query=$([uri]::EscapeDataString("$b $suf"))&display=1" $null $null).total
        $subRows += [pscustomobject]@{
            brand = $b; topic = $t; blog_query = "$b $suf"
            search_index_12m = [math]::Round((Expand-Series $res $yPeriods | Measure-Object ratio -Sum).Sum * $factor, 1)
            blog_posts = [int64]$total
        }
    }
    Write-Host "  - call $call : $(($topics[$i..($i + 1)] | ForEach-Object name) -join ', ')"
}
Save-Csv $subRows 'topics_12m.csv'

Write-Host '5) 블로그·뉴스 글 수집 (정확도순·최신순 각 100건)'
$themeDefs = [ordered]@{
    blog = [ordered]@{
        '착용 후기'       = '후기|리뷰|착용|내돈내산|구매'
        '사이즈·핏'       = '사이즈|핏|치수|기장'
        '레깅스'          = '레깅스'
        '요가·필라테스'   = '요가|필라테스'
        '러닝·헬스'       = '러닝|런닝|헬스|운동복|짐웨어'
        '골프'            = '골프'
        '남성'            = '남자|남성|맨즈|남편'
        '수영·물놀이'     = '수영|스윔|래쉬가드|비키니'
        '할인·쿠폰'       = '할인|세일|쿠폰|특가|적립|블프|블랙프라이데이'
        '일상복·출근룩'   = '데일리|일상|출근|코디|원마일'
        '비교'            = '비교|차이|vs|VS'
        '협찬 표시'       = '협찬|제공받|원고료|소정의|체험단'
    }
    news = [ordered]@{
        '실적·매출'       = '매출|영업이익|실적|흑자|적자|분기|순이익|성장률'
        '신제품·컬렉션'   = '출시|신제품|컬렉션|론칭|런칭|선보'
        '모델·광고'       = '모델|앰버서더|엠버서더|광고|캠페인|화보'
        '해외 진출'       = '해외|글로벌|일본|중국|대만|미국|싱가포르|진출|수출'
        '매장·팝업'       = '매장|팝업|오프라인|스토어|백화점|플래그십'
        '할인·프로모션'   = '할인|세일|프로모션|특가|쿠폰|블랙프라이데이'
        '협업·후원'       = '협업|콜라보|후원|파트너|제휴|공식'
        '주가·투자'       = '주가|투자|상장|주식|목표가|배당|증권'
        'ESG·사회공헌'    = '기부|ESG|친환경|사회공헌|리사이클|재활용'
        '골프'            = '골프'
        '러닝·요가 등 종목' = '러닝|마라톤|요가|필라테스|테니스|수영'
    }
}
$stop = @('있는','있습니다','합니다','하는','그리고','하지만','정말','너무','이번','오늘','많이','같은','위한','대한','통해','에서','으로','입니다','있어요','하고','해서','안녕하세요','여러분','그래서','이렇게','어떻게','진짜','먼저','바로','우리','지금','젝시믹스','안다르','xexymix','andar','XEXYMIX','ANDAR','있다','했다','밝혔다','이다','것으로','등을','따르면','지난','함께','또한','특히','이후','관련','gt','lt','quot','amp')

$items = @(); $volRows = @()
$queries = @()
foreach ($b in $brands.Keys) { foreach ($ch in 'blog', 'news') { foreach ($so in 'sim', 'date') { $queries += @{ brand = $b; q = $b; ch = $ch; sort = $so; match = $brands[$b].match; file = $brands[$b].file } } } }
$queries += @{ brand = '두 브랜드 함께'; q = '젝시믹스 안다르'; ch = 'blog'; sort = 'sim'; match = "(?=.*($($brands['젝시믹스'].match)))(?=.*($($brands['안다르'].match)))"; file = 'both' }

foreach ($q in $queries) {
    $saveAs = "{0}_{1}_{2}.json" -f $q.ch, $q.file, $q.sort
    try {
        $r = Invoke-NaverApi 'GET' "/search/v1/$($q.ch)?query=$([uri]::EscapeDataString($q.q))&display=100&sort=$($q.sort)" $null $saveAs
    } catch {
        $code = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { '' }
        Write-Host "  - $($q.brand) $($q.ch) $($q.sort): 건너뜀 ($code)"
        $volRows += [pscustomobject]@{ brand = $q.brand; channel = $q.ch; sort = $q.sort; total = "오류 $code" }
        continue
    }
    $dates = @()
    foreach ($it in $r.items) {
        $title = Clear-Html $it.title; $desc = Clear-Html $it.description
        $date = if ($it.postdate) { [datetime]::ParseExact($it.postdate, 'yyyyMMdd', $null) } else { [datetime]$it.pubDate }
        $link = if ($it.originallink) { $it.originallink } else { $it.link }
        $source = if ($it.bloggername) { $it.bloggername } else { try { ([uri]$link).Host -replace '^www\.', '' } catch { '' } }
        $relevant = "$title $desc" -match $q.match
        $themes = @($themeDefs[$q.ch].Keys | Where-Object { "$title $desc" -match $themeDefs[$q.ch][$_] })
        $dates += $date
        $items += [pscustomobject]@{ brand = $q.brand; channel = $q.ch; sort = $q.sort; date = $date.ToString('yyyy-MM-dd'); source = $source
            relevant = $relevant; themes = ($themes -join '; '); title = $title; description = $desc; link = $link }
    }
    $relCount = @($items | Where-Object { $_.brand -eq $q.brand -and $_.channel -eq $q.ch -and $_.sort -eq $q.sort -and $_.relevant }).Count
    $span = if ($dates.Count) { [math]::Max(1, (($dates | Measure-Object -Maximum).Maximum - ($dates | Measure-Object -Minimum).Minimum).TotalDays) } else { 0 }
    $volRows += [pscustomobject]@{ brand = $q.brand; channel = $q.ch; sort = $q.sort; total = [int64]$r.total; fetched = $dates.Count; relevant = $relCount
        oldest = $(if ($dates) { ($dates | Measure-Object -Minimum).Minimum.ToString('yyyy-MM-dd') }); newest = $(if ($dates) { ($dates | Measure-Object -Maximum).Maximum.ToString('yyyy-MM-dd') })
        relevant_per_week = $(if ($q.sort -eq 'date' -and $span) { [math]::Round($relCount / $span * 7, 1) } else { '' }) }
    Write-Host "  - $($q.brand) $($q.ch) $($q.sort): 전체 $($r.total)건 중 $($dates.Count)건 수집, 관련 $($relCount)건"
}
Save-Csv $items 'content_items.csv'
Save-Csv $volRows 'content_volume.csv'

# 같은 글이 정확도순·최신순에 모두 나올 수 있으므로 링크 기준으로 한 번만 셈
$unique = $items | Where-Object relevant | Group-Object brand, channel, link | ForEach-Object { $_.Group[0] }
$themeRows = foreach ($g in ($unique | Group-Object brand, channel)) {
    $first = $g.Group[0]
    foreach ($t in $themeDefs[$first.channel].Keys) {
        $c = @($g.Group | Where-Object { ($_.themes -split '; ') -contains $t }).Count
        [pscustomobject]@{ brand = $first.brand; channel = $first.channel; theme = $t; posts = $c; base = $g.Count; pct = [math]::Round($c / $g.Count * 100, 1) }
    }
}
Save-Csv $themeRows 'content_themes.csv'

# 제목·요약에서 자주 나오는 표현 (형태소 분석 없는 단순 빈도)
$wordRows = foreach ($g in ($unique | Group-Object brand, channel)) {
    $words = (($g.Group | ForEach-Object { "$($_.title) $($_.description)" }) -join ' ') -split '[^\p{L}\p{Nd}]+' |
        ForEach-Object { if ($_.Length -ge 3) { $_ -replace '(을|를|은|는|이|가|의|에|도|과|와|로|으로|에서)$', '' } else { $_ } } |
        Where-Object { $_.Length -ge 2 -and $stop -notcontains $_ -and $_ -notmatch '^\d+$' } |
        Group-Object | Sort-Object Count -Descending | Select-Object -First 40
    $rank = 0
    foreach ($w in $words) { $rank++; [pscustomobject]@{ brand = $g.Group[0].brand; channel = $g.Group[0].channel; rank = $rank; word = $w.Name; count = $w.Count } }
}
Save-Csv $wordRows 'content_words.csv'

# 뉴스 출처(언론사 도메인) 상위
$pressRows = foreach ($g in ($unique | Where-Object channel -eq 'news' | Group-Object brand)) {
    $g.Group | Group-Object source | Sort-Object Count -Descending | Select-Object -First 10 |
        ForEach-Object { [pscustomobject]@{ brand = $g.Name; source = $_.Name; articles = $_.Count } }
}
Save-Csv $pressRows 'news_sources.csv'

$info = [ordered]@{ collectedAt = $now.ToString('yyyy-MM-dd HH:mm'); monthly = "$(Fmt $mStart) ~ $(Fmt $lastMonthEnd)"
    weekly = "$(Fmt $weekStart) ~ $(Fmt $lastSunday)"; last12m = "$(Fmt $y1Start) ~ $(Fmt $lastMonthEnd)" }
[IO.File]::WriteAllText((Join-Path $outDir 'collection_info.json'), ($info | ConvertTo-Json), $utf8)

Write-Host "`n완료: $outDir"
