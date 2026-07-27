Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$publisherPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\Publish-QPayBackend.ps1'))
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('QPayBackendPublisherTests-[QPay]-' + [Guid]::NewGuid().ToString('N'))
$testFailed = $false

function Assert-True {
    param([bool]$Condition, [string]$Message)

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)

    if ($Expected -ne $Actual) {
        throw $Message
    }
}

if (-not (Test-Path -LiteralPath $publisherPath -PathType Leaf)) {
    Write-Error 'Publisher script is missing: Publish-QPayBackend.ps1'
    exit 1
}

New-Item -ItemType Directory -Path $temporaryRoot | Out-Null

try {
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $publisherOutput = & powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass `
            -File $publisherPath `
            -Configuration Release `
            -ArtifactsRoot $temporaryRoot 2>&1 | Out-String
        $publisherExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    Assert-Equal 0 $publisherExitCode 'Official publisher should exit successfully.'

    $publishedDirectories = @(Get-ChildItem -LiteralPath $temporaryRoot -Directory -Filter 'QPayBackend-Release-*')
    $zipFiles = @(Get-ChildItem -LiteralPath $temporaryRoot -File -Filter 'QPayBackend-Release-*.zip')
    $checksumFiles = @(Get-ChildItem -LiteralPath $temporaryRoot -File -Filter 'QPayBackend-Release-*.zip.sha256')

    Assert-Equal 1 $publishedDirectories.Count 'Publisher should create exactly one fresh deployment directory.'
    Assert-Equal 1 $zipFiles.Count 'Publisher should create exactly one ZIP archive.'
    Assert-Equal 1 $checksumFiles.Count 'Publisher should create exactly one ZIP checksum file.'

    $publishedDirectory = $publishedDirectories[0].FullName
    $runtimeConfig = Join-Path $publishedDirectory 'QPayBackend.exe.config'
    $forbiddenConfig = Join-Path $publishedDirectory 'app.config'
    $manifestPath = Join-Path $publishedDirectory 'deployment-manifest.json'

    Assert-True (Test-Path -LiteralPath $runtimeConfig -PathType Leaf) 'Runtime config is missing from the deployment directory.'
    Assert-True (-not (Test-Path -LiteralPath $forbiddenConfig)) 'Source app.config must not be in the deployment directory.'
    Assert-True (Test-Path -LiteralPath $manifestPath -PathType Leaf) 'Deployment manifest is missing.'

    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
    $runtimeConfigHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $runtimeConfig).Hash
    $zipHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zipFiles[0].FullName).Hash
    $checksumText = Get-Content -Raw -LiteralPath $checksumFiles[0].FullName

    Assert-Equal 'QPayBackend' $manifest.application 'Manifest application is incorrect.'
    Assert-Equal 'Release' $manifest.configuration 'Manifest configuration is incorrect.'
    Assert-Equal 'QPayBackend.exe.config' $manifest.runtimeConfig 'Manifest runtime config filename is incorrect.'
    Assert-Equal $runtimeConfigHash $manifest.runtimeConfigSha256 'Manifest runtime config checksum is incorrect.'
    Assert-True ($checksumText.IndexOf($zipHash, [StringComparison]::Ordinal) -ge 0) 'ZIP checksum file does not contain the ZIP SHA256.'
    Assert-True ($publisherOutput.IndexOf('[QPayPublish] publication-succeeded', [StringComparison]::Ordinal) -ge 0) 'Publisher did not report a success marker.'

    Write-Host 'PASS: official publisher creates a clean, verified deployment package.'
}
catch {
    $testFailed = $true
    Write-Host ('FAIL: official publisher integration ({0})' -f $_.Exception.Message)
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
