[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceConfig,

    [Parameter(Mandatory = $true)]
    [string]$RuntimeConfig,

    [Parameter(Mandatory = $true)]
    [string]$ForbiddenConfig
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$failures = New-Object 'System.Collections.Generic.List[string]'

function Add-ValidationFailure {
    param(
        [string]$Category,
        [string]$Key
    )

    $message = if ([string]::IsNullOrEmpty($Key)) {
        '[QPayConfig] {0}' -f $Category
    }
    else {
        '[QPayConfig] {0} key={1}' -f $Category, $Key
    }

    if (-not $failures.Contains($message)) {
        $failures.Add($message)
    }
}

function Get-AppSettingsMap {
    param(
        [string]$Path,
        [ValidateSet('source', 'runtime')]
        [string]$Kind
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Add-ValidationFailure -Category ('missing-{0}-config' -f $Kind)
        return $null
    }

    try {
        [xml]$document = Get-Content -Raw -LiteralPath $Path
    }
    catch {
        Add-ValidationFailure -Category ('invalid-{0}-xml' -f $Kind)
        return $null
    }

    $settings = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    $nodes = @($document.SelectNodes('/configuration/appSettings/add'))
    if ($nodes.Count -eq 0) {
        Add-ValidationFailure -Category 'missing-appsettings' -Key $Kind
        return $null
    }

    foreach ($node in $nodes) {
        $key = [string]$node.key
        if ([string]::IsNullOrEmpty($key)) {
            Add-ValidationFailure -Category 'empty-key' -Key $Kind
            continue
        }

        if ($settings.ContainsKey($key)) {
            Add-ValidationFailure -Category 'duplicate-key' -Key $key
            continue
        }

        $settings.Add($key, [string]$node.value)
    }

    return $settings
}

function Test-IsSensitiveQPaySetting {
    param([string]$Key)

    return [string]::Equals($Key, 'SINOPAC_SITE', [StringComparison]::Ordinal) -or
        [string]::Equals($Key, 'SANDBOX_SITE', [StringComparison]::Ordinal) -or
        $Key.EndsWith('_XKeyID', [StringComparison]::Ordinal)
}

function Test-SafeSettingValue {
    param(
        [string]$Key,
        [string]$Value
    )

    if (-not (Test-IsSensitiveQPaySetting -Key $Key)) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($Value)) {
        Add-ValidationFailure -Category 'unsafe-empty-value' -Key $Key
        return
    }

    if (-not [string]::Equals($Value, $Value.Trim(), [StringComparison]::Ordinal)) {
        Add-ValidationFailure -Category 'unsafe-outer-whitespace' -Key $Key
    }
}

try {
    if (Test-Path -LiteralPath $ForbiddenConfig) {
        Add-ValidationFailure -Category 'forbidden-published-file' -Key 'app.config'
    }

    $sourceSettings = Get-AppSettingsMap -Path $SourceConfig -Kind source
    $runtimeSettings = Get-AppSettingsMap -Path $RuntimeConfig -Kind runtime

    if ($null -ne $sourceSettings) {
        foreach ($key in $sourceSettings.Keys) {
            Test-SafeSettingValue -Key $key -Value $sourceSettings[$key]
        }
    }

    if ($null -ne $runtimeSettings) {
        foreach ($key in $runtimeSettings.Keys) {
            Test-SafeSettingValue -Key $key -Value $runtimeSettings[$key]
        }
    }

    if (($null -ne $sourceSettings) -and ($null -ne $runtimeSettings)) {
        foreach ($key in $sourceSettings.Keys) {
            if (-not $runtimeSettings.ContainsKey($key)) {
                Add-ValidationFailure -Category 'missing-runtime-key' -Key $key
                continue
            }

            if (-not [string]::Equals($sourceSettings[$key], $runtimeSettings[$key], [StringComparison]::Ordinal)) {
                Add-ValidationFailure -Category 'value-mismatch' -Key $key
            }
        }

        foreach ($key in $runtimeSettings.Keys) {
            if (-not $sourceSettings.ContainsKey($key)) {
                Add-ValidationFailure -Category 'unexpected-runtime-key' -Key $key
            }
        }
    }
}
catch {
    Add-ValidationFailure -Category 'verifier-internal-error'
}

if ($failures.Count -gt 0) {
    foreach ($failure in $failures) {
        [Console]::Error.WriteLine($failure)
    }
    exit 1
}

Write-Host ('[QPayConfig] validation-succeeded settings={0}' -f $sourceSettings.Count)
exit 0
