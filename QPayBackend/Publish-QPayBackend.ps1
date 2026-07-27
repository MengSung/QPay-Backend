[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [string]$ArtifactsRoot
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$stage = 'initialize'
$publishDirectory = $null
$resolvedArtifactsRoot = $null
$zipPath = $null
$checksumPath = $null
$temporaryZipPath = $null
$temporaryChecksumPath = $null

function Test-PathsEqual {
    param([string]$Left, [string]$Right)

    return [string]::Equals(
        $Left.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar),
        $Right.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar),
        [StringComparison]::OrdinalIgnoreCase)
}

function Remove-SafePublicationPath {
    param(
        [string]$Path,
        [switch]$Recurse
    )

    if ([string]::IsNullOrEmpty($Path) -or
        [string]::IsNullOrEmpty($resolvedArtifactsRoot) -or
        -not (Test-Path -LiteralPath $Path)) {
        return
    }

    try {
        $resolvedPath = [IO.Path]::GetFullPath($Path)
        $rootPrefix = $resolvedArtifactsRoot.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
        if (-not $resolvedPath.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            return
        }

        if ($Recurse) {
            Remove-Item -LiteralPath $resolvedPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        else {
            Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        return
    }
}

function Remove-PartialPublication {
    Remove-SafePublicationPath -Path $publishDirectory -Recurse
    Remove-SafePublicationPath -Path $temporaryZipPath
    Remove-SafePublicationPath -Path $temporaryChecksumPath
    Remove-SafePublicationPath -Path $checksumPath
    Remove-SafePublicationPath -Path $zipPath
}

try {
    $projectDirectory = [IO.Path]::GetFullPath($PSScriptRoot)
    $repositoryRoot = [IO.Path]::GetFullPath((Join-Path $projectDirectory '..'))
    $projectPath = Join-Path $projectDirectory 'QPayBackend.csproj'

    if ([string]::IsNullOrWhiteSpace($ArtifactsRoot)) {
        $ArtifactsRoot = Join-Path $repositoryRoot 'artifacts'
    }

    $resolvedArtifactsRoot = [IO.Path]::GetFullPath($ArtifactsRoot)
    $driveRoot = [IO.Path]::GetPathRoot($resolvedArtifactsRoot)
    if ([string]::IsNullOrWhiteSpace($resolvedArtifactsRoot) -or
        (Test-PathsEqual -Left $resolvedArtifactsRoot -Right $driveRoot) -or
        (Test-PathsEqual -Left $resolvedArtifactsRoot -Right $repositoryRoot) -or
        (Test-PathsEqual -Left $resolvedArtifactsRoot -Right $projectDirectory)) {
        throw 'unsafe-artifacts-root'
    }

    $stage = 'create-output'
    New-Item -ItemType Directory -Force -Path $resolvedArtifactsRoot | Out-Null
    $artifactName = 'QPayBackend-{0}-{1}-{2}' -f `
        $Configuration,
        (Get-Date -Format 'yyyyMMdd-HHmmss'),
        [Guid]::NewGuid().ToString('N').Substring(0, 8)
    $publishDirectory = Join-Path $resolvedArtifactsRoot $artifactName
    New-Item -ItemType Directory -Path $publishDirectory | Out-Null

    $stage = 'clean'
    & dotnet clean $projectPath -c $Configuration --nologo
    if ($LASTEXITCODE -ne 0) {
        throw 'dotnet-clean-failed'
    }

    $stage = 'publish'
    & dotnet publish $projectPath -c $Configuration --output $publishDirectory --nologo
    if ($LASTEXITCODE -ne 0) {
        throw 'dotnet-publish-failed'
    }

    $stage = 'verify-output'
    $runtimeConfigPath = Join-Path $publishDirectory 'QPayBackend.exe.config'
    $forbiddenConfigPath = Join-Path $publishDirectory 'app.config'
    if (-not (Test-Path -LiteralPath $runtimeConfigPath -PathType Leaf)) {
        throw 'runtime-config-missing'
    }
    if (Test-Path -LiteralPath $forbiddenConfigPath) {
        throw 'source-config-published'
    }

    $stage = 'write-manifest'
    $runtimeConfigHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $runtimeConfigPath).Hash
    $manifest = [ordered]@{
        schemaVersion = 1
        application = 'QPayBackend'
        configuration = $Configuration
        createdAtUtc = [DateTime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        runtimeConfig = 'QPayBackend.exe.config'
        runtimeConfigSha256 = $runtimeConfigHash
    }
    $manifestPath = Join-Path $publishDirectory 'deployment-manifest.json'
    $manifestJson = $manifest | ConvertTo-Json
    [IO.File]::WriteAllText($manifestPath, $manifestJson + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))

    $stage = 'create-archive'
    $zipPath = $publishDirectory + '.zip'
    $checksumPath = $zipPath + '.sha256'
    $temporaryZipPath = $zipPath + '.partial'
    $temporaryChecksumPath = $checksumPath + '.partial'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::CreateFromDirectory(
        $publishDirectory,
        $temporaryZipPath,
        [IO.Compression.CompressionLevel]::Optimal,
        $false)
    $zipHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $temporaryZipPath).Hash
    $checksumText = '{0}  {1}{2}' -f $zipHash, [IO.Path]::GetFileName($zipPath), [Environment]::NewLine
    [IO.File]::WriteAllText($temporaryChecksumPath, $checksumText, (New-Object Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $temporaryChecksumPath -Destination $checksumPath
    Move-Item -LiteralPath $temporaryZipPath -Destination $zipPath

    Write-Host '[QPayPublish] publication-succeeded'
    Write-Host ('[QPayPublish] deployment-directory={0}' -f $publishDirectory)
    Write-Host ('[QPayPublish] archive={0}' -f $zipPath)
    Write-Host ('[QPayPublish] checksum={0}' -f $checksumPath)
    exit 0
}
catch {
    try {
        Remove-PartialPublication
    }
    catch {
    }
    [Console]::Error.WriteLine(('[QPayPublish] publication-failed stage={0}' -f $stage))
    exit 1
}
