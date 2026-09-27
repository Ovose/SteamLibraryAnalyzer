
param(
    [string]$SteamApiKey = "",
    [string]$OutputPath = "$PSScriptRoot\Steam_Analysis.xlsx"
)

$ErrorActionPreference = "Stop"

# эта штука считает мертвыми игры где вообще 0 минут
$DeadGameMaxMinutes = 0

# это вроде нормально работает для заброшенных игр
$AbandonedMinHours = 20
$AbandonedDays = 365

# без задержки стим иногда начинает ругаться
$StoreDelayMs = 300

$CountryCode = "US"

$FranchisePatterns = [ordered]@{
    "STAR WARS"          = "(?i)\bSTAR\s*WARS\b"
    "Call of Duty"       = "(?i)\bCall of Duty\b|\bCOD\b"
    "Need for Speed"     = "(?i)\bNeed for Speed\b"
    "Grand Theft Auto"   = "(?i)\bGrand Theft Auto\b|\bGTA\b"
    "Assassin's Creed"   = "(?i)\bAssassin'?s Creed\b"
    "Resident Evil"      = "(?i)\bResident Evil\b"
    "Half-Life"          = "(?i)\bHalf[- ]Life\b"
    "Portal"             = "(?i)\bPortal\b"
    "Fallout"            = "(?i)\bFallout\b"
    "The Elder Scrolls"  = "(?i)\bThe Elder Scrolls\b|\bSkyrim\b"
    "Far Cry"            = "(?i)\bFar Cry\b"
    "Battlefield"        = "(?i)\bBattlefield\b"
    "Forza"              = "(?i)\bForza\b"
    "DOOM"               = "(?i)\bDOOM\b"
    "Wolfenstein"        = "(?i)\bWolfenstein\b"
    "Metro"              = "(?i)^Metro\b|\bMetro 2033\b|\bMetro Exodus\b"
    "Tom Clancy"         = "(?i)\bTom Clancy'?s\b"
    "Counter-Strike"     = "(?i)\bCounter-Strike\b"
}

# тут просто ввод ссылок, особо трогать не надо
function Read-ProfileLinks {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host " STEAM LIBRARY ANALYZER" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "IMPORTANT:" -ForegroundColor Yellow
    Write-Host "The Steam profile must be PUBLIC." -ForegroundColor Yellow
    Write-Host "Game details must also be PUBLIC." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Steam privacy path:" -ForegroundColor DarkGray
    Write-Host "Profile -> Edit Profile -> Privacy Settings" -ForegroundColor DarkGray
    Write-Host "My profile = Public" -ForegroundColor DarkGray
    Write-Host "Game details = Public" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Accepted examples:" -ForegroundColor DarkGray
    Write-Host "https://steamcommunity.com/id/ExampleName/" -ForegroundColor DarkGray
    Write-Host "https://steamcommunity.com/profiles/7656119XXXXXXXXXX/" -ForegroundColor DarkGray
    Write-Host ""

    $items = @()

    while ($items.Count -lt 2) {
        $number = $items.Count + 1
        if ($number -eq 1) {
            $url = Read-Host "Steam profile link #1"
        }
        else {
            $url = Read-Host "Steam profile link #2 (press Enter to skip)"
            if ([string]::IsNullOrWhiteSpace($url)) {
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace($url)) {
            Write-Host "Profile #1 is required." -ForegroundColor Red
            continue
        }

        $url = $url.Trim()

        if ($url -notmatch '^https?://steamcommunity\.com/(id|profiles)/') {
            Write-Host "This does not look like a Steam profile link. Try again." -ForegroundColor Red
            continue
        }

        $items += [PSCustomObject]@{
            Name = "Profile$number"
            Url  = $url
        }
    }

    return @($items)
}

$Profiles = @(Read-ProfileLinks)

# старый вариант был с профилями прямо в коде, оставил на всякий случай
# $Profiles = @(
#     [PSCustomObject]@{ Name = "test1"; Url = "https://steamcommunity.com/id/example1/" },
#     [PSCustomObject]@{ Name = "test2"; Url = "https://steamcommunity.com/id/example2/" }
# )

$CachePath = "$PSScriptRoot\steam_store_cache.json"

# если на новом пк нет модуля, он сам попробует поставить
function Ensure-ImportExcel {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return
    }

    Write-Host ""
    Write-Host "ImportExcel module is not installed." -ForegroundColor Yellow
    Write-Host "Installing it automatically for the current Windows user..." -ForegroundColor Yellow

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nuget) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -Confirm:$false | Out-Null
        }

        $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($gallery -and $gallery.InstallationPolicy -ne "Trusted") {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }

        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop

        Write-Host "ImportExcel installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "Automatic ImportExcel installation failed." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host ""
        Write-Host "Try running this manually once:" -ForegroundColor Yellow
        Write-Host "Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Cyan
        throw
    }
}

function Get-ApiKey {
    param([string]$CurrentKey)

    if (-not [string]::IsNullOrWhiteSpace($CurrentKey)) {
        return $CurrentKey.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($env:STEAM_API_KEY)) {
        return $env:STEAM_API_KEY.Trim()
    }

    Write-Host ""
    Write-Host "Enter Steam Web API key:" -ForegroundColor Cyan
    $key = Read-Host "Steam API key"

    if ([string]::IsNullOrWhiteSpace($key)) {
        throw "Steam API key was not entered."
    }

    return $key.Trim()
}

function Invoke-SteamApi {
    param([string]$Uri)
    try {
        return Invoke-RestMethod -Uri $Uri -Method Get -TimeoutSec 45 -Headers @{
            "User-Agent" = "Mozilla/5.0 SteamLibraryAnalyzer"
        }
    }
    catch {
        Write-Warning "Request failed: $Uri"
        Write-Warning $_.Exception.Message
        return $null
    }
}

function Resolve-SteamId {
    param(
        [string]$ProfileUrl,
        [string]$ApiKey
    )

    if ($ProfileUrl -match "/profiles/(\d{17})") {
        return $Matches[1]
    }

    if ($ProfileUrl -match "/id/([^/?#]+)") {
        $vanity = [uri]::EscapeDataString($Matches[1])
        $uri = "https://api.steampowered.com/ISteamUser/ResolveVanityURL/v1/?key=$ApiKey&vanityurl=$vanity"
        $r = Invoke-SteamApi $uri
        if ($r -and $r.response.success -eq 1) {
            return [string]$r.response.steamid
        }
    }

    return $null
}

function Get-PlayerSummary {
    param(
        [string]$SteamId,
        [string]$ApiKey
    )
    $uri = "https://api.steampowered.com/ISteamUser/GetPlayerSummaries/v2/?key=$ApiKey&steamids=$SteamId"
    $r = Invoke-SteamApi $uri
    if ($r -and $r.response.players.Count -gt 0) {
        return $r.response.players[0]
    }
    return $null
}

function Get-OwnedGames {
    param(
        [string]$SteamId,
        [string]$ApiKey
    )
    $uri = "https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/?key=$ApiKey&steamid=$SteamId&include_appinfo=true&include_played_free_games=true&include_free_sub=true"
    $r = Invoke-SteamApi $uri
    if ($r -and $r.response.games) {
        return @($r.response.games)
    }
    return @()
}

function Get-RecentlyPlayedGames {
    param(
        [string]$SteamId,
        [string]$ApiKey
    )
    $uri = "https://api.steampowered.com/IPlayerService/GetRecentlyPlayedGames/v1/?key=$ApiKey&steamid=$SteamId&count=100"
    $r = Invoke-SteamApi $uri
    if ($r -and $r.response.games) {
        return @($r.response.games)
    }
    return @()
}

function Load-StoreCache {
    $cache = @{}
    if (Test-Path $CachePath) {
        try {
            $obj = Get-Content $CachePath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) {
                $cache[$p.Name] = $p.Value
            }
        }
        catch {
            Write-Warning "Could not read store cache. A new cache will be created."
        }
    }
    return $cache
}

function Save-StoreCache {
    param([hashtable]$Cache)
    $Cache | ConvertTo-Json -Depth 12 | Set-Content -Path $CachePath -Encoding UTF8
}

# отсюда тянутся цены и жанры, иногда у удаленных игр пусто
function Get-StoreInfo {
    param(
        [int]$AppId,
        [hashtable]$Cache
    )

    $key = [string]$AppId
    if ($Cache.ContainsKey($key)) {
        return $Cache[$key]
    }

    $uri = "https://store.steampowered.com/api/appdetails?appids=$AppId&cc=$CountryCode&l=english"
    $r = Invoke-SteamApi $uri

    $info = [PSCustomObject]@{
        AppId        = $AppId
        StoreSuccess = $false
        Type         = ""
        IsFree       = $false
        PriceUSD     = $null
        InitialUSD   = $null
        DiscountPct  = $null
        Genres       = @()
        Developers   = @()
        Publishers   = @()
        ReleaseDate  = ""
    }

    try {
        $entry = $r.$key
        if ($entry -and $entry.success -eq $true -and $entry.data) {
            $d = $entry.data
            $info.StoreSuccess = $true
            $info.Type = [string]$d.type
            $info.IsFree = [bool]$d.is_free

            if ($d.price_overview) {
                $info.PriceUSD = [math]::Round(([double]$d.price_overview.final / 100), 2)
                $info.InitialUSD = [math]::Round(([double]$d.price_overview.initial / 100), 2)
                $info.DiscountPct = [int]$d.price_overview.discount_percent
            }
            elseif ($info.IsFree) {
                $info.PriceUSD = 0.0
                $info.InitialUSD = 0.0
                $info.DiscountPct = 0
            }

            if ($d.genres) {
                $info.Genres = @($d.genres | ForEach-Object { [string]$_.description })
            }
            if ($d.developers) {
                $info.Developers = @($d.developers | ForEach-Object { [string]$_ })
            }
            if ($d.publishers) {
                $info.Publishers = @($d.publishers | ForEach-Object { [string]$_ })
            }
            if ($d.release_date) {
                $info.ReleaseDate = [string]$d.release_date.date
            }
        }
    }
    catch {
    }

    $Cache[$key] = $info
    Start-Sleep -Milliseconds $StoreDelayMs
    return $info
}

# костыль для серий, но пока работает
function Find-Franchise {
    param([string]$GameName)

    $franchiseMatches = @()

    foreach ($kv in $FranchisePatterns.GetEnumerator()) {
        if ($GameName -match $kv.Value) {
            $franchiseMatches += [string]$kv.Key
        }
    }

    return @($franchiseMatches)
}

function Format-Hours {
    param([double]$Minutes)
    return [math]::Round($Minutes / 60.0, 1)
}

function Get-PortraitText {
    param(
        [object]$Stats,
        [object[]]$GenreRows,
        [object[]]$FranchiseRows
    )

    $parts = New-Object System.Collections.Generic.List[string]

    if ($Stats.TotalGames -gt 0) {
        $playedPct = [math]::Round((($Stats.TotalGames - $Stats.NeverPlayed) * 100.0 / $Stats.TotalGames), 0)
        $parts.Add("В библиотеке $($Stats.TotalGames) игр; запущено хотя бы раз примерно $playedPct%.")
    }

    if ($GenreRows.Count -gt 0) {
        $top = @($GenreRows | Sort-Object GameCount -Descending | Select-Object -First 3)
        if ($top.Count -gt 0) {
            $genreText = ($top | ForEach-Object { "$($_.Genre) ($($_.GameCount))" }) -join ", "
            $parts.Add("Самые заметные жанры: $genreText.")
        }
    }

    if ($FranchiseRows.Count -gt 0) {
        $topF = @($FranchiseRows | Sort-Object TotalHours -Descending | Select-Object -First 3)
        if ($topF.Count -gt 0) {
            $frText = ($topF | ForEach-Object { "$($_.Franchise) ($($_.TotalHours) ч.)" }) -join ", "
            $parts.Add("По времени сильнее всего выделяются серии: $frText.")
        }
    }

    if ($Stats.Hours100Plus -gt 0) {
        $parts.Add("Игр с 100+ часами: $($Stats.Hours100Plus).")
    }

    if ($Stats.NeverPlayed -gt 0) {
        $deadPct = [math]::Round(($Stats.NeverPlayed * 100.0 / [math]::Max(1,$Stats.TotalGames)), 0)
        $parts.Add("Ни разу не запускались $($Stats.NeverPlayed) игр ($deadPct% библиотеки).")
    }

    if ($Stats.AbandonedCount -gt 0) {
        $parts.Add("Заброшенных игр по текущему правилу: $($Stats.AbandonedCount).")
    }

    if ($Stats.Hours2Weeks -gt 0) {
        $parts.Add("За последние 2 недели: $($Stats.Hours2Weeks) ч.")
    }
    else {
        $parts.Add("За последние 2 недели публичная активность не найдена.")
    }

    return ($parts -join " ")
}

Ensure-ImportExcel
$SteamApiKey = Get-ApiKey $SteamApiKey
# кэш ускоряет второй запуск, удалять можно если что-то совсем странное
$storeCache = Load-StoreCache

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

$allProfiles = @()
$allGamesRows = @()
$allGenreRows = @()
$allFranchiseRows = @()
$allDeadRows = @()
$allAbandonedRows = @()
$profileGameMaps = @{}

foreach ($profile in $Profiles) {

    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor DarkCyan
    Write-Host "Analyzing $($profile.Name)" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor DarkCyan

    $steamId = Resolve-SteamId -ProfileUrl $profile.Url -ApiKey $SteamApiKey
    if (-not $steamId) {
        Write-Warning "Could not resolve SteamID for $($profile.Url)"
        continue
    }

    $summary = Get-PlayerSummary -SteamId $steamId -ApiKey $SteamApiKey
    if ($summary -and -not [string]::IsNullOrWhiteSpace([string]$summary.personaname)) {
        $profile.Name = [string]$summary.personaname
    }
    $games = @(Get-OwnedGames -SteamId $steamId -ApiKey $SteamApiKey)
    $recent = @(Get-RecentlyPlayedGames -SteamId $steamId -ApiKey $SteamApiKey)

    if ($games.Count -eq 0) {
        Write-Warning "No public library returned for $($profile.Name). Is Game details public?"
        continue
    }

    $recent2WeekMap = @{}
    foreach ($g in $recent) {
        $recent2WeekMap[[string]$g.appid] = if ($g.playtime_2weeks) { [int]$g.playtime_2weeks } else { 0 }
    }

    $gameMap = @{}
    $genreCounter = @{}
    $genreHours = @{}
    $franchiseCounter = @{}
    $franchiseHours = @{}

    $profileGamesRows = @()
    $profileDeadRows = @()
    $profileAbandonedRows = @()

    $i = 0
    foreach ($g in $games) {
        $i++
        Write-Progress -Activity "Steam Store metadata: $($profile.Name)" -Status "$i / $($games.Count): $($g.name)" -PercentComplete (($i * 100.0) / $games.Count)

        $store = Get-StoreInfo -AppId ([int]$g.appid) -Cache $storeCache
        $hours = Format-Hours ([double]$g.playtime_forever)
        $hours2w = if ($recent2WeekMap.ContainsKey([string]$g.appid)) { Format-Hours ([double]$recent2WeekMap[[string]$g.appid]) } else { 0.0 }

        $lastPlayed = $null
        if ($g.rtime_last_played -and [int64]$g.rtime_last_played -gt 0) {
            $lastPlayed = [DateTimeOffset]::FromUnixTimeSeconds([int64]$g.rtime_last_played).LocalDateTime
        }

        $daysSince = $null
        if ($lastPlayed) {
            $daysSince = [int][math]::Floor(((Get-Date) - $lastPlayed).TotalDays)
        }

        $genres = @($store.Genres)
        foreach ($genre in $genres) {
            if (-not $genreCounter.ContainsKey($genre)) {
                $genreCounter[$genre] = 0
                $genreHours[$genre] = 0.0
            }
            $genreCounter[$genre]++
            $genreHours[$genre] += $hours
        }

        $franchises = @(Find-Franchise -GameName ([string]$g.name))
        foreach ($fr in $franchises) {
            if (-not $franchiseCounter.ContainsKey($fr)) {
                $franchiseCounter[$fr] = 0
                $franchiseHours[$fr] = 0.0
            }
            $franchiseCounter[$fr]++
            $franchiseHours[$fr] += $hours
        }

        $isDead = ([int]$g.playtime_forever -le $DeadGameMaxMinutes)
        $isAbandoned = ($hours -ge $AbandonedMinHours -and $daysSince -ne $null -and $daysSince -ge $AbandonedDays)

        $row = [PSCustomObject]@{
            Profile         = $profile.Name
            AppID           = [int]$g.appid
            Game            = [string]$g.name
            HoursTotal      = $hours
            Hours2Weeks     = $hours2w
            LastPlayed      = $lastPlayed
            DaysSincePlayed = $daysSince
            Dead            = $isDead
            Abandoned       = $isAbandoned
            Genres          = ($genres -join ", ")
            Franchises      = ($franchises -join ", ")
            IsFree          = [bool]$store.IsFree
            CurrentPriceUSD = $store.PriceUSD
            ListPriceUSD    = $store.InitialUSD
            DiscountPct     = $store.DiscountPct
            StoreType       = [string]$store.Type
            ReleaseDate     = [string]$store.ReleaseDate
        }

        $profileGamesRows += $row
        $allGamesRows += $row
        $gameMap[[string]$g.appid] = $row

        if ($isDead) {
            $deadRow = [PSCustomObject]@{
                Profile = $profile.Name
                AppID   = [int]$g.appid
                Game    = [string]$g.name
                Hours   = $hours
            }
            $profileDeadRows += $deadRow
            $allDeadRows += $deadRow
        }

        if ($isAbandoned) {
            $abRow = [PSCustomObject]@{
                Profile         = $profile.Name
                AppID           = [int]$g.appid
                Game            = [string]$g.name
                Hours           = $hours
                LastPlayed      = $lastPlayed
                DaysSincePlayed = $daysSince
            }
            $profileAbandonedRows += $abRow
            $allAbandonedRows += $abRow
        }
    }

    Write-Progress -Activity "Steam Store metadata: $($profile.Name)" -Completed
    Save-StoreCache $storeCache

    $profileGenreRows = @()
    foreach ($genre in $genreCounter.Keys) {
        $r = [PSCustomObject]@{
            Profile    = $profile.Name
            Genre      = $genre
            GameCount  = [int]$genreCounter[$genre]
            TotalHours = [math]::Round([double]$genreHours[$genre], 1)
        }
        $profileGenreRows += $r
        $allGenreRows += $r
    }

    $profileFranchiseRows = @()
    foreach ($fr in $franchiseCounter.Keys) {
        $r = [PSCustomObject]@{
            Profile    = $profile.Name
            Franchise  = $fr
            GameCount  = [int]$franchiseCounter[$fr]
            TotalHours = [math]::Round([double]$franchiseHours[$fr], 1)
        }
        $profileFranchiseRows += $r
        $allFranchiseRows += $r
    }

    $neverPlayed = @($profileGamesRows | Where-Object { $_.HoursTotal -eq 0 }).Count
    $under1 = @($profileGamesRows | Where-Object { $_.HoursTotal -gt 0 -and $_.HoursTotal -lt 1 }).Count
    $h1to10 = @($profileGamesRows | Where-Object { $_.HoursTotal -ge 1 -and $_.HoursTotal -lt 10 }).Count
    $h10to50 = @($profileGamesRows | Where-Object { $_.HoursTotal -ge 10 -and $_.HoursTotal -lt 50 }).Count
    $h50to100 = @($profileGamesRows | Where-Object { $_.HoursTotal -ge 50 -and $_.HoursTotal -lt 100 }).Count
    $h100plus = @($profileGamesRows | Where-Object { $_.HoursTotal -ge 100 }).Count

    $hours2Weeks = [math]::Round((($profileGamesRows | Measure-Object Hours2Weeks -Sum).Sum), 1)
    $totalHours = [math]::Round((($profileGamesRows | Measure-Object HoursTotal -Sum).Sum), 1)

    $priced = @($profileGamesRows | Where-Object { $_.CurrentPriceUSD -ne $null })
    $pricedList = @($profileGamesRows | Where-Object { $_.ListPriceUSD -ne $null })

    $currentValue = if ($priced.Count -gt 0) { [math]::Round((($priced | Measure-Object CurrentPriceUSD -Sum).Sum), 2) } else { 0 }
    $listValue = if ($pricedList.Count -gt 0) { [math]::Round((($pricedList | Measure-Object ListPriceUSD -Sum).Sum), 2) } else { 0 }

    $stats = [PSCustomObject]@{
        Profile             = $profile.Name
        SteamID             = $steamId
        ProfileURL          = $profile.Url
        PersonaName         = if ($summary) { $summary.personaname } else { $profile.Name }
        TotalGames          = $profileGamesRows.Count
        TotalHours          = $totalHours
        Hours2Weeks         = $hours2Weeks
        NeverPlayed         = $neverPlayed
        Under1Hour          = $under1
        Hours1to10          = $h1to10
        Hours10to50         = $h10to50
        Hours50to100        = $h50to100
        Hours100Plus        = $h100plus
        AbandonedCount      = $profileAbandonedRows.Count
        CurrentValueUSD     = $currentValue
        ListValueUSD        = $listValue
        GamesWithPrice      = $priced.Count
        Portrait            = ""
    }

    $stats.Portrait = Get-PortraitText -Stats $stats -GenreRows $profileGenreRows -FranchiseRows $profileFranchiseRows

    $allProfiles += $stats
    $profileGameMaps[$profile.Name] = $gameMap
}

Save-StoreCache $storeCache

if ($allProfiles.Count -eq 0) {
    throw "No profile data could be collected."
}

$compatibilityRows = @()
$onlyRows = @()

if ($allProfiles.Count -ge 2) {
    $p1 = $allProfiles[0].Profile
    $p2 = $allProfiles[1].Profile
    $m1 = $profileGameMaps[$p1]
    $m2 = $profileGameMaps[$p2]

    $ids1 = @($m1.Keys)
    $ids2 = @($m2.Keys)

    $commonIds = @($ids1 | Where-Object { $m2.ContainsKey($_) })
    $only1 = @($ids1 | Where-Object { -not $m2.ContainsKey($_) })
    $only2 = @($ids2 | Where-Object { -not $m1.ContainsKey($_) })

    $unionCount = @($ids1 + $ids2 | Sort-Object -Unique).Count
    $compatPct = if ($unionCount -gt 0) { [math]::Round(($commonIds.Count * 100.0 / $unionCount), 1) } else { 0 }

    $compatibilityRows += [PSCustomObject]@{
        Profile1       = $p1
        Profile2       = $p2
        GamesProfile1  = $ids1.Count
        GamesProfile2  = $ids2.Count
        CommonGames    = $commonIds.Count
        UnionGames     = $unionCount
        CompatibilityPct = $compatPct
    }

    foreach ($id in $commonIds) {
        $g1 = $m1[$id]
        $g2 = $m2[$id]
        $onlyRows += [PSCustomObject]@{
            Type       = "Общая игра"
            Game       = $g1.Game
            AppID      = [int]$id
            Profile    = "$p1 + $p2"
            Hours      = [math]::Round(($g1.HoursTotal + $g2.HoursTotal), 1)
        }
    }

    foreach ($id in $only1) {
        $g = $m1[$id]
        $onlyRows += [PSCustomObject]@{
            Type       = "Только у $p1"
            Game       = $g.Game
            AppID      = [int]$id
            Profile    = $p1
            Hours      = $g.HoursTotal
        }
    }

    foreach ($id in $only2) {
        $g = $m2[$id]
        $onlyRows += [PSCustomObject]@{
            Type       = "Только у $p2"
            Game       = $g.Game
            AppID      = [int]$id
            Profile    = $p2
            Hours      = $g.HoursTotal
        }
    }
}

$structureRows = @()
foreach ($p in $allProfiles) {
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="0 ч. / не запускалась"; Games=$p.NeverPlayed }
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="0–1 ч."; Games=$p.Under1Hour }
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="1–10 ч."; Games=$p.Hours1to10 }
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="10–50 ч."; Games=$p.Hours10to50 }
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="50–100 ч."; Games=$p.Hours50to100 }
    $structureRows += [PSCustomObject]@{ Profile=$p.Profile; Bucket="100+ ч."; Games=$p.Hours100Plus }
}

$excelParams = @{
    Path          = $OutputPath
    AutoSize      = $true
    AutoFilter    = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableStyle    = "Medium2"
}

$allProfiles |
    Select-Object Profile, PersonaName, ProfileURL, SteamID, TotalGames, TotalHours, Hours2Weeks,
        NeverPlayed, Under1Hour, Hours1to10, Hours10to50, Hours50to100, Hours100Plus,
        AbandonedCount, CurrentValueUSD, ListValueUSD, GamesWithPrice, Portrait |
    Export-Excel @excelParams -WorksheetName "Сводка"

$allGamesRows |
    Sort-Object Profile, @{Expression="HoursTotal"; Descending=$true} |
    Export-Excel -Path $OutputPath -WorksheetName "Все игры" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

$allDeadRows |
    Sort-Object Profile, Game |
    Export-Excel -Path $OutputPath -WorksheetName "Мертвые игры" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

$allAbandonedRows |
    Sort-Object Profile, @{Expression="DaysSincePlayed"; Descending=$true} |
    Export-Excel -Path $OutputPath -WorksheetName "Заброшенные" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

$allGenreRows |
    Sort-Object Profile, @{Expression="GameCount"; Descending=$true}, @{Expression="TotalHours"; Descending=$true} |
    Export-Excel -Path $OutputPath -WorksheetName "Жанры" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

$allFranchiseRows |
    Sort-Object Profile, @{Expression="TotalHours"; Descending=$true}, @{Expression="GameCount"; Descending=$true} |
    Export-Excel -Path $OutputPath -WorksheetName "Серии" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

$structureRows |
    Export-Excel -Path $OutputPath -WorksheetName "Структура" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append

if ($compatibilityRows.Count -gt 0) {
    $compatibilityRows |
        Export-Excel -Path $OutputPath -WorksheetName "Совместимость" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append
}

if ($onlyRows.Count -gt 0) {
    $onlyRows |
        Sort-Object Type, Game |
        Export-Excel -Path $OutputPath -WorksheetName "Общие и различия" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -Append
}


# названия в самом экселе, а внутри кода пусть остаются английские чтобы ничего не ломать
$RussianHeaders = @{
    "Profile"          = "Профиль"
    "PersonaName"      = "Имя в Steam"
    "ProfileURL"       = "Ссылка на профиль"
    "SteamID"          = "Steam ID"
    "TotalGames"       = "Всего игр"
    "TotalHours"       = "Всего часов"
    "Hours2Weeks"      = "Часов за 2 недели"
    "NeverPlayed"      = "Не запускались"
    "Under1Hour"       = "Меньше 1 часа"
    "Hours1to10"       = "От 1 до 10 часов"
    "Hours10to50"      = "От 10 до 50 часов"
    "Hours50to100"     = "От 50 до 100 часов"
    "Hours100Plus"     = "100+ часов"
    "AbandonedCount"   = "Заброшенных игр"
    "CurrentValueUSD"  = "Текущая стоимость, $"
    "ListValueUSD"     = "Стоимость без скидок, $"
    "GamesWithPrice"   = "Игр с известной ценой"
    "Portrait"         = "Портрет игрока"

    "AppID"            = "App ID"
    "Game"             = "Игра"
    "HoursTotal"       = "Всего часов"
    "LastPlayed"       = "Последний запуск"
    "DaysSincePlayed"  = "Дней с запуска"
    "Dead"             = "Мёртвая"
    "Abandoned"        = "Заброшенная"
    "Genres"           = "Жанры"
    "Franchises"       = "Серии"
    "IsFree"           = "Бесплатная"
    "CurrentPriceUSD"  = "Текущая цена, $"
    "ListPriceUSD"     = "Цена без скидки, $"
    "DiscountPct"      = "Скидка, %"
    "StoreType"        = "Тип в магазине"
    "ReleaseDate"      = "Дата выхода"

    "Hours"            = "Часы"
    "Genre"            = "Жанр"
    "GameCount"        = "Количество игр"
    "Franchise"        = "Игровая серия"

    "Bucket"           = "Диапазон часов"
    "Games"            = "Количество игр"

    "Profile1"         = "Профиль 1"
    "Profile2"         = "Профиль 2"
    "GamesProfile1"    = "Игр у профиля 1"
    "GamesProfile2"    = "Игр у профиля 2"
    "CommonGames"      = "Общих игр"
    "UnionGames"       = "Уникальных игр всего"
    "CompatibilityPct" = "Совпадение библиотек, %"

    "Type"             = "Тип"
}

$pkg = Open-ExcelPackage -Path $OutputPath
try {
    foreach ($ws in $pkg.Workbook.Worksheets) {
        if (-not $ws.Dimension) { continue }

        $ws.View.FreezePanes(2, 1)

        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            $oldHeader = [string]$ws.Cells[1,$c].Value
            if ($RussianHeaders.ContainsKey($oldHeader)) {
                $ws.Cells[1,$c].Value = $RussianHeaders[$oldHeader]
            }
        }

        $header = $ws.Cells[1,1,1,$ws.Dimension.End.Column]
        $header.Style.Font.Bold = $true
        $header.Style.WrapText = $true

        $ws.Cells[$ws.Dimension.Address].Style.VerticalAlignment = "Top"
        $ws.Cells[$ws.Dimension.Address].AutoFitColumns()

        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            if ($ws.Column($c).Width -gt 45) {
                $ws.Column($c).Width = 45
                $ws.Column($c).Style.WrapText = $true
            }
        }
    }

    $ws = $pkg.Workbook.Worksheets["Сводка"]
    if ($ws) {
        $ws.Column(15).Style.Numberformat.Format = '$0.00'
        $ws.Column(16).Style.Numberformat.Format = '$0.00'
        $ws.Column(18).Width = 70
        $ws.Column(18).Style.WrapText = $true
    }

    $ws = $pkg.Workbook.Worksheets["Все игры"]
    if ($ws) {
        $ws.Column(6).Style.Numberformat.Format = "yyyy-mm-dd hh:mm"
        $ws.Column(13).Style.Numberformat.Format = '$0.00'
        $ws.Column(14).Style.Numberformat.Format = '$0.00'
    }

    $ws = $pkg.Workbook.Worksheets["Заброшенные"]
    if ($ws) {
        $ws.Column(5).Style.Numberformat.Format = "yyyy-mm-dd"
    }

    Close-ExcelPackage $pkg
}
catch {
    Close-ExcelPackage $pkg -NoSave
    throw
}

# это хотел сделать для красивого графика прямо в экселе, но пока не работает как надо
# $chart = New-ExcelChartDefinition -XRange "Жанры!B2:B20" -YRange "Жанры!C2:C20"
# Export-Excel -Path $OutputPath -WorksheetName "Жанры" -ExcelChartDefinition $chart -Append

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Green
Write-Host "DONE" -ForegroundColor Green
Write-Host "Excel file:" -ForegroundColor Green
Write-Host $OutputPath -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Green
Write-Host ""
Write-Host "Sheets created:" -ForegroundColor Cyan
Write-Host "  Сводка"
Write-Host "  Все игры"
Write-Host "  Мертвые игры"
Write-Host "  Заброшенные"
Write-Host "  Жанры"
Write-Host "  Серии"
Write-Host "  Структура"
Write-Host "  Совместимость"
Write-Host "  Общие и различия"

if ($allProfiles.Count -lt 2) {
    Write-Host ""
    Write-Host "Compatibility sheets are skipped because only one profile was entered." -ForegroundColor DarkYellow
}
