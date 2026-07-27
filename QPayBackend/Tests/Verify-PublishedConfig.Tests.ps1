Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$verifierPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\Verify-PublishedConfig.ps1'))
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('QPayConfigVerifierTests-' + [Guid]::NewGuid().ToString('N'))

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)

    if ($Expected -ne $Actual) {
        throw $Message
    }
}

function Assert-Contains {
    param([string]$Text, [string]$ExpectedText, [string]$Message)

    if ($Text.IndexOf($ExpectedText, [StringComparison]::Ordinal) -lt 0) {
        throw $Message
    }
}

function Assert-NotContains {
    param([string]$Text, [string]$ForbiddenText, [string]$Message)

    if ($Text.IndexOf($ForbiddenText, [StringComparison]::Ordinal) -ge 0) {
        throw $Message
    }
}

function Write-AppSettingsConfig {
    param(
        [string]$Path,
        [System.Collections.IDictionary]$Settings,
        [switch]$IncludeStartup
    )

    $builder = New-Object Text.StringBuilder
    [void]$builder.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
    [void]$builder.AppendLine('<configuration>')
    [void]$builder.AppendLine('  <appSettings>')

    foreach ($entry in $Settings.GetEnumerator()) {
        $escapedKey = [Security.SecurityElement]::Escape([string]$entry.Key)
        $escapedValue = [Security.SecurityElement]::Escape([string]$entry.Value)
        [void]$builder.AppendLine(('    <add key="{0}" value="{1}" />' -f $escapedKey, $escapedValue))
    }

    [void]$builder.AppendLine('  </appSettings>')
    if ($IncludeStartup) {
        [void]$builder.AppendLine('  <startup><supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.1" /></startup>')
    }
    [void]$builder.AppendLine('</configuration>')

    [IO.File]::WriteAllText($Path, $builder.ToString(), (New-Object Text.UTF8Encoding($false)))
}

function Invoke-Verifier {
    param(
        [string]$SourceConfig,
        [string]$RuntimeConfig,
        [string]$ForbiddenConfig
    )

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass `
            -File $verifierPath `
            -SourceConfig $SourceConfig `
            -RuntimeConfig $RuntimeConfig `
            -ForbiddenConfig $ForbiddenConfig 2>&1 | Out-String
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    [pscustomobject]@{
        ExitCode = $exitCode
        Output = $output
    }
}

function Invoke-Test {
    param([string]$Name, [scriptblock]$Body)

    try {
        & $Body
        $script:Passed += 1
        Write-Host ('PASS: {0}' -f $Name)
    }
    catch {
        $script:Failed += 1
        Write-Host ('FAIL: {0} ({1})' -f $Name, $_.Exception.Message)
    }
}

if (-not (Test-Path -LiteralPath $verifierPath -PathType Leaf)) {
    Write-Error 'Verifier script is missing: Verify-PublishedConfig.ps1'
    exit 1
}

New-Item -ItemType Directory -Path $temporaryRoot | Out-Null

try {
    Invoke-Test 'matching settings succeed' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'matching')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $settings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $settings
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $settings -IncludeStartup
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 0 $result.ExitCode 'Matching settings should succeed.'
    }

    Invoke-Test 'changed value fails without value disclosure' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'mismatch')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $sourceSecret = 'SOURCE-611-XKEY-SENTINEL'
        $runtimeSecret = 'RUNTIME-OLD-XKEY-SENTINEL'
        $sourceSettings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = $sourceSecret
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }
        $runtimeSettings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = $runtimeSecret
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $sourceSettings
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $runtimeSettings -IncludeStartup
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Changed value should fail.'
        Assert-Contains $result.Output '611_XKeyID' 'Failure should identify the changed key.'
        Assert-NotContains $result.Output $sourceSecret 'Failure disclosed the source value.'
        Assert-NotContains $result.Output $runtimeSecret 'Failure disclosed the runtime value.'
    }

    Invoke-Test 'missing runtime key fails' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'missing-key')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $sourceSettings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }
        $runtimeSettings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $sourceSettings
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $runtimeSettings -IncludeStartup
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Missing runtime key should fail.'
        Assert-Contains $result.Output 'NANKAN_XKeyID' 'Failure should identify the missing key.'
    }

    Invoke-Test 'outer whitespace fails' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'whitespace')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $settings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = ' PADDED-XKEY-SENTINEL '
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $settings
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $settings -IncludeStartup
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Outer whitespace should fail.'
        Assert-Contains $result.Output 'unsafe-outer-whitespace' 'Failure should report the whitespace category.'
        Assert-NotContains $result.Output 'PADDED-XKEY-SENTINEL' 'Failure disclosed the padded value.'
    }

    Invoke-Test 'missing runtime config fails' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'missing-runtime')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $settings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $settings
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Missing runtime config should fail.'
        Assert-Contains $result.Output 'missing-runtime-config' 'Failure should report the missing runtime config.'
    }

    Invoke-Test 'published app.config fails' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'forbidden')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $settings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        Write-AppSettingsConfig -Path $sourceConfig -Settings $settings
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $settings -IncludeStartup
        Write-AppSettingsConfig -Path $forbiddenConfig -Settings $settings
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Published app.config should fail.'
        Assert-Contains $result.Output 'forbidden-published-file' 'Failure should report the forbidden file.'
    }

    Invoke-Test 'missing appSettings section fails clearly' {
        $caseDirectory = New-Item -ItemType Directory -Path (Join-Path $temporaryRoot 'missing-appsettings')
        $sourceConfig = Join-Path $caseDirectory.FullName 'source.config'
        $runtimeConfig = Join-Path $caseDirectory.FullName 'runtime.config'
        $forbiddenConfig = Join-Path $caseDirectory.FullName 'app.config'
        $settings = [ordered]@{
            SINOPAC_SITE = 'https://example.invalid/qpay/'
            SANDBOX_SITE = 'https://sandbox.example.invalid/qpay/'
            '611_XKeyID' = 'MATCHING-611-XKEY'
            NANKAN_XKeyID = 'MATCHING-NANKAN-XKEY'
        }

        [IO.File]::WriteAllText(
            $sourceConfig,
            '<?xml version="1.0" encoding="utf-8"?><configuration><startup /></configuration>',
            (New-Object Text.UTF8Encoding($false)))
        Write-AppSettingsConfig -Path $runtimeConfig -Settings $settings -IncludeStartup
        $result = Invoke-Verifier -SourceConfig $sourceConfig -RuntimeConfig $runtimeConfig -ForbiddenConfig $forbiddenConfig

        Assert-Equal 1 $result.ExitCode 'Missing appSettings should fail.'
        Assert-Contains $result.Output 'missing-appsettings' 'Failure should report the missing appSettings section.'
        Assert-Contains $result.Output 'source' 'Failure should identify the source document.'
        Assert-NotContains $result.Output 'verifier-internal-error' 'Missing appSettings should not become an internal error.'
    }
}
finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $resolvedSystemTemporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolvedTemporaryRoot.StartsWith($resolvedSystemTemporaryRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedTemporaryRoot)) {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
}

Write-Host ('RESULT: {0} passed, {1} failed' -f $script:Passed, $script:Failed)
if ($script:Failed -gt 0) {
    exit 1
}

exit 0
