<#
  .env 파일에서 API 키를 읽어 오는 공통 도구 (다른 스크립트가 불러서 씀)

  사용 예 (다른 .ps1 안에서):
    . (Join-Path $root 'tools\load-env.ps1')
    $config = Get-NaverConfig $root     # clientId / clientSecret / baseUrl
    $envs   = Import-DotEnv $root       # .env 전체 (예: $envs.KAKAO_REST_API_KEY)

  키 위치: 프로젝트 폴더의 .env (GitHub 업로드 제외)
#>

function Import-DotEnv([string]$Root) {
    $path = Join-Path $Root '.env'
    if (-not (Test-Path $path)) { throw ".env 파일이 없습니다: $path" }
    $vars = @{}
    foreach ($line in [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)) {
        if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$') {
            $vars[$Matches[1]] = $Matches[2].Trim('"', "'")
        }
    }
    $vars
}

function Get-NaverConfig([string]$Root) {
    $vars = Import-DotEnv $Root
    foreach ($k in 'NAVER_CLIENT_ID', 'NAVER_CLIENT_SECRET') {
        if (-not $vars[$k]) { throw ".env에 $k 값이 없습니다." }
    }
    $base = if ($vars['NAVER_BASE_URL']) { $vars['NAVER_BASE_URL'] } else { 'https://naverapihub.apigw.ntruss.com' }
    [pscustomobject]@{ clientId = $vars['NAVER_CLIENT_ID']; clientSecret = $vars['NAVER_CLIENT_SECRET']; baseUrl = $base }
}
