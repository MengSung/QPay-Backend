Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$projectPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\QPayBackend.csproj'))
$temporaryDirectory = Join-Path ([IO.Path]::GetTempPath()) ('QPayBackendPublishTests-' + [Guid]::NewGuid().ToString('N'))
$testFailed = $false

New-Item -ItemType Directory -Path $temporaryDirectory | Out-Null

try {
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $publishOutput = & dotnet publish $projectPath -c Release --output $temporaryDirectory --nologo 2>&1 | Out-String
        $publishExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($publishExitCode -ne 0) {
        throw ('dotnet publish failed with exit code {0}.' -f $publishExitCode)
    }

    $runtimeConfig = Join-Path $temporaryDirectory 'QPayBackend.exe.config'
    $forbiddenConfig = Join-Path $temporaryDirectory 'app.config'

    if (-not (Test-Path -LiteralPath $runtimeConfig -PathType Leaf)) {
        throw 'QPayBackend.exe.config was not generated.'
    }

    if (Test-Path -LiteralPath $forbiddenConfig) {
        throw 'Published app.config must be absent.'
    }

    if ($publishOutput.IndexOf('[QPayConfig] validation-succeeded', [StringComparison]::Ordinal) -lt 0) {
        throw 'Project publish did not execute the configuration verifier.'
    }

    Write-Host 'PASS: publish generates only the effective runtime config.'
}
catch {
    $testFailed = $true
    Write-Host ('FAIL: publish configuration integration ({0})' -f $_.Exception.Message)
}
finally {
    $resolvedTemporaryDirectory = [IO.Path]::GetFullPath($temporaryDirectory)
    $resolvedSystemTemporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolvedTemporaryDirectory.StartsWith($resolvedSystemTemporaryRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $resolvedTemporaryDirectory)) {
        Remove-Item -LiteralPath $resolvedTemporaryDirectory -Recurse -Force
    }
}

if ($testFailed) {
    exit 1
}

exit 0
