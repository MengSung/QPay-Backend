Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\QPayBackend.csproj'))
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('QPayBackendPublishTests-' + [Guid]::NewGuid().ToString('N'))
$successfulPublishDirectory = Join-Path $temporaryRoot 'matching'
$stalePublishDirectory = Join-Path $temporaryRoot 'stale'
$injectionScriptPath = Join-Path $temporaryRoot 'Inject-StaleConfig.ps1'
$injectionTargetsPath = Join-Path $temporaryRoot 'Inject-StaleConfig.targets'
$testFailed = $false

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)

    if ($Expected -ne $Actual) {
        throw $Message
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)

    if (-not $Condition) {
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

function Invoke-ProjectPublish {
    param(
        [string]$OutputDirectory,
        [string]$CustomTargetsPath
    )

    $arguments = @(
        'publish',
        $projectPath,
        '-c',
        'Release',
        '--output',
        $OutputDirectory,
        '--nologo'
    )
    if (-not [string]::IsNullOrWhiteSpace($CustomTargetsPath)) {
        $arguments += ('-p:CustomAfterMicrosoftCommonTargets={0}' -f $CustomTargetsPath)
    }

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & dotnet @arguments 2>&1 | Out-String
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

function Write-StaleConfigInjection {
    $injectionScript = @'
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeConfig
)

$ErrorActionPreference = 'Stop'
[xml]$document = Get-Content -Raw -LiteralPath $RuntimeConfig
$setting = $document.SelectSingleNode("/configuration/appSettings/add[@key='611_XKeyID']")
if ($null -eq $setting) {
    exit 2
}

$setting.value = 'INTEGRATION-STALE-XKEY-SENTINEL'
$document.Save($RuntimeConfig)
exit 0
'@
    [IO.File]::WriteAllText($injectionScriptPath, $injectionScript, (New-Object Text.UTF8Encoding($false)))

    $injectionTargets = @'
<Project>
  <Target Name="InjectStalePublishedConfigForTest" BeforeTargets="ValidatePublishedConfiguration">
    <Exec Command="powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File &quot;$(MSBuildThisFileDirectory)Inject-StaleConfig.ps1&quot; -RuntimeConfig &quot;$(PublishDir)$(TargetFileName).config&quot;" />
  </Target>
</Project>
'@
    [IO.File]::WriteAllText($injectionTargetsPath, $injectionTargets, (New-Object Text.UTF8Encoding($false)))
}

New-Item -ItemType Directory -Path $temporaryRoot | Out-Null

try {
    $successfulPublish = Invoke-ProjectPublish -OutputDirectory $successfulPublishDirectory
    Assert-Equal 0 $successfulPublish.ExitCode 'Matching project publish should succeed.'

    $runtimeConfig = Join-Path $successfulPublishDirectory 'QPayBackend.exe.config'
    $forbiddenConfig = Join-Path $successfulPublishDirectory 'app.config'
    Assert-True (Test-Path -LiteralPath $runtimeConfig -PathType Leaf) 'QPayBackend.exe.config was not generated.'
    Assert-True (-not (Test-Path -LiteralPath $forbiddenConfig)) 'Published app.config must be absent.'
    Assert-Contains $successfulPublish.Output '[QPayConfig] validation-succeeded' 'Project publish did not execute the configuration verifier.'
    Write-Host 'PASS: publish generates only the effective runtime config.'

    Write-StaleConfigInjection
    $stalePublish = Invoke-ProjectPublish -OutputDirectory $stalePublishDirectory -CustomTargetsPath $injectionTargetsPath
    Assert-True ($stalePublish.ExitCode -ne 0) 'Stale runtime config must abort dotnet publish.'
    Assert-Contains $stalePublish.Output '[QPayConfig] value-mismatch key=611_XKeyID' 'Failed publish should identify only the mismatched key.'
    Assert-NotContains $stalePublish.Output 'INTEGRATION-STALE-XKEY-SENTINEL' 'Failed publish disclosed the injected setting value.'
    Assert-NotContains $stalePublish.Output '[QPayConfig] validation-succeeded' 'Failed publish must not report validation success.'
    Write-Host 'PASS: stale runtime config aborts the MSBuild publish target.'
}
catch {
    $testFailed = $true
    Write-Host ('FAIL: publish configuration integration ({0})' -f $_.Exception.Message)
}
finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $resolvedSystemTemporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolvedTemporaryRoot.StartsWith($resolvedSystemTemporaryRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedTemporaryRoot)) {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
}

if ($testFailed) {
    exit 1
}

exit 0
