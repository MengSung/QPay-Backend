# QPay Runtime Config Publish Guard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prevent deployment of stale QPay XKeys by verifying generated runtime settings and removing misleading `app.config` from every deployment package.

**Architecture:** Keep `QPayBackend/app.config` as the only source configuration, but never publish that source filename. A project-level publish target invokes a value-redacting PowerShell verifier against generated `QPayBackend.exe.config`. A clean publisher always uses a fresh output directory and emits a manifest, ZIP, and checksum only after validation succeeds.

**Tech Stack:** ASP.NET Core 2.1 on .NET Framework 4.7.1, SDK-style MSBuild, Windows PowerShell 5.1, `dotnet publish`, Visual Studio MSBuild.

---

### Task 1: Build the redacting configuration verifier with TDD

**Files:**

- Create: `QPayBackend/Tests/Verify-PublishedConfig.Tests.ps1`
- Create: `QPayBackend/Verify-PublishedConfig.ps1`

- [ ] **Step 1: Write the failing verifier test harness**

Create a dependency-free PowerShell test harness that writes temporary source/runtime XML fixtures containing sentinel values, invokes the real verifier in a child `powershell.exe`, and checks exit code plus combined output.

The harness must contain these six cases:

```powershell
Invoke-Test 'matching settings succeed' { Assert-ExitCode 0 }
Invoke-Test 'changed value fails without value disclosure' { Assert-ExitCode 1; Assert-Contains '611_XKeyID'; Assert-NotContains $sourceSecret; Assert-NotContains $runtimeSecret }
Invoke-Test 'missing runtime key fails' { Assert-ExitCode 1; Assert-Contains 'NANKAN_XKeyID' }
Invoke-Test 'outer whitespace fails' { Assert-ExitCode 1; Assert-Contains 'unsafe-outer-whitespace' }
Invoke-Test 'missing runtime config fails' { Assert-ExitCode 1; Assert-Contains 'missing-runtime-config' }
Invoke-Test 'published app.config fails' { Assert-ExitCode 1; Assert-Contains 'forbidden-published-file' }
```

All fixtures use fake sentinel values; no repository XKey is read or copied into test output.

- [ ] **Step 2: Run the verifier tests and verify RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\QPayBackend\Tests\Verify-PublishedConfig.Tests.ps1
```

Expected: nonzero exit with `Verifier script is missing`, proving the required implementation does not yet exist.

- [ ] **Step 3: Implement the minimal verifier**

`Verify-PublishedConfig.ps1` must declare:

```powershell
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SourceConfig,
    [Parameter(Mandatory = $true)][string]$RuntimeConfig,
    [Parameter(Mandatory = $true)][string]$ForbiddenConfig
)
```

Implementation requirements:

```powershell
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

# Parse only configuration/appSettings/add into an ordinal dictionary.
# Compare source/runtime keys and values with StringComparison.Ordinal.
# Validate SINOPAC_SITE, SANDBOX_SITE, and keys ending in _XKeyID.
# Add only fixed categories and key names to a failure list.
# Never write a value, complete XML, or raw exception message.
# Exit 1 when any failure exists; otherwise print one success summary and exit 0.
```

Required fixed categories are `missing-source-config`, `missing-runtime-config`, `invalid-source-xml`, `invalid-runtime-xml`, `duplicate-key`, `missing-runtime-key`, `unexpected-runtime-key`, `value-mismatch`, `unsafe-empty-value`, `unsafe-outer-whitespace`, `forbidden-published-file`, and `verifier-internal-error`.

- [ ] **Step 4: Run verifier tests and verify GREEN**

Run the command from Step 2.

Expected: `6 passed, 0 failed`, exit code 0, and no sentinel setting value in output.

### Task 2: Make every project publish enforce the guard

**Files:**

- Create: `QPayBackend/Tests/Publish-Configuration.Tests.ps1`
- Modify: `QPayBackend/QPayBackend.csproj`

- [ ] **Step 1: Write the failing publish integration test**

The script publishes to a GUID-named directory beneath the system temporary directory and asserts:

```powershell
dotnet publish $projectPath -c Release --output $publishDirectory --nologo
Assert-True (Test-Path (Join-Path $publishDirectory 'QPayBackend.exe.config'))
Assert-False (Test-Path (Join-Path $publishDirectory 'app.config'))
```

The `finally` block may recursively delete only the exact GUID directory it created beneath `[IO.Path]::GetTempPath()`.

- [ ] **Step 2: Run publish integration test and verify RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\QPayBackend\Tests\Publish-Configuration.Tests.ps1
```

Expected: failure because current Web SDK publication includes `app.config`.

- [ ] **Step 3: Wire source configuration and validation into MSBuild**

Add to the main property group in `QPayBackend.csproj`:

```xml
<AppConfig>$(MSBuildProjectDirectory)\app.config</AppConfig>
```

Add:

```xml
<ItemGroup>
  <Content Update="app.config">
    <CopyToOutputDirectory>Never</CopyToOutputDirectory>
    <CopyToPublishDirectory>Never</CopyToPublishDirectory>
  </Content>
</ItemGroup>

<Target Name="ValidatePublishedConfiguration" AfterTargets="Publish">
  <Exec
    WorkingDirectory="$(MSBuildProjectDirectory)"
    Command="powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File &quot;$(MSBuildProjectDirectory)\Verify-PublishedConfig.ps1&quot; -SourceConfig &quot;$(AppConfig)&quot; -RuntimeConfig &quot;$(PublishDir)$(TargetFileName).config&quot; -ForbiddenConfig &quot;$(PublishDir)app.config&quot;" />
</Target>
```

- [ ] **Step 4: Run publish integration test and verify GREEN**

Run the command from Step 2.

Expected: publish exits 0, `QPayBackend.exe.config` exists, `app.config` is absent, and the verifier prints a success summary.

### Task 3: Replace the stale reusable output workflow

**Files:**

- Create: `QPayBackend/Tests/Publish-QPayBackend.Tests.ps1`
- Create: `QPayBackend/Publish-QPayBackend.ps1`
- Modify: `QPayBackend/DotNetPublish-Release.bat`

- [ ] **Step 1: Write the failing publisher integration test**

The test invokes the official publisher with a GUID-named temporary artifact root and asserts:

```powershell
powershell.exe -File $publisher -Configuration Release -ArtifactsRoot $temporaryRoot
Assert-ExitCode 0
Assert-Count 1 (Get-ChildItem $temporaryRoot -Directory -Filter 'QPayBackend-Release-*')
Assert-Count 1 (Get-ChildItem $temporaryRoot -File -Filter 'QPayBackend-Release-*.zip')
Assert-Count 1 (Get-ChildItem $temporaryRoot -File -Filter 'QPayBackend-Release-*.zip.sha256')
Assert-False (Test-Path (Join-Path $publishedDirectory 'app.config'))
Assert-True (Test-Path (Join-Path $publishedDirectory 'QPayBackend.exe.config'))
Assert-True (Test-Path (Join-Path $publishedDirectory 'deployment-manifest.json'))
```

- [ ] **Step 2: Run publisher test and verify RED**

Run:

```powershell
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File .\QPayBackend\Tests\Publish-QPayBackend.Tests.ps1
```

Expected: nonzero exit because `Publish-QPayBackend.ps1` does not yet exist.

- [ ] **Step 3: Implement the clean publisher**

`Publish-QPayBackend.ps1` parameters:

```powershell
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Release',
    [string]$ArtifactsRoot
)
```

The implementation must:

```powershell
# Default ArtifactsRoot to <repo>/artifacts.
# Reject a drive root, repository root, or empty resolved path.
# Create QPayBackend-<Configuration>-yyyyMMdd-HHmmss-<8-char-guid> as a new directory.
# Run dotnet clean and dotnet publish against QPayBackend.csproj.
# Stop immediately when either command returns nonzero.
# Confirm runtime config exists and source app.config is absent.
# Write deployment-manifest.json containing application, configuration,
# createdAtUtc, runtimeConfig filename, and runtimeConfigSha256 only.
# Compress the directory into a sibling ZIP and write <zip>.sha256.
# On failure, print only a fixed stage name and exit nonzero.
```

- [ ] **Step 4: Replace the Release batch entry point**

Use a thin wrapper that preserves the PowerShell exit code:

```bat
@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Publish-QPayBackend.ps1" -Configuration Release
set "QPAY_PUBLISH_EXIT=%ERRORLEVEL%"
if not "%QPAY_PUBLISH_EXIT%"=="0" echo [QPayPublish] FAILED. Nothing is ready for deployment.
if "%QPAY_PUBLISH_EXIT%"=="0" echo [QPayPublish] SUCCESS.
pause
exit /b %QPAY_PUBLISH_EXIT%
```

- [ ] **Step 5: Run publisher test and verify GREEN**

Run the command from Step 2.

Expected: one fresh deployment directory, one ZIP, one checksum file, valid manifest, runtime config present, source config absent, and exit code 0.

### Task 4: Verify complete change and record review

**Files:**

- Modify: `.ccg/tasks/fix-shekinah611-payment-postprocessing/task.json`
- Create: `.ccg/tasks/fix-shekinah611-payment-postprocessing/review.md`

- [ ] **Step 1: Run all PowerShell tests**

```powershell
Get-ChildItem .\QPayBackend\Tests\*.Tests.ps1 | ForEach-Object {
    & powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $_.FullName
    if ($LASTEXITCODE -ne 0) { throw "Test failed: $($_.Name)" }
}
```

Expected: every script exits 0.

- [ ] **Step 2: Run full solution build**

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' .\QPayBackend.sln /m /t:Build /p:Configuration=Release /v:minimal
```

Expected: exit code 0. Existing warnings must be recorded but no new compiler error is allowed.

- [ ] **Step 3: Inspect scope and secret safety**

```powershell
git diff --check
git diff --stat
git status --short
rg -n "MyPay" QPayBackend\Verify-PublishedConfig.ps1 QPayBackend\Publish-QPayBackend.ps1 QPayBackend\Tests
```

Expected: only publish guard, tests, docs, and CCG files changed; no MyPay production file changed; no real XKey appears in new test or documentation files.

- [ ] **Step 4: Perform review and fix Critical findings**

Attempt the required Gemini and Claude review in parallel with a bounded wait. If external review times out, stop it under the user's explicit authorization and perform a local plus `ccg-review` review. Record Critical/Warning/Info findings and resolutions in `review.md`.

- [ ] **Step 5: Complete and archive the CCG task**

Set task status/phase to completed, record the verified commands, move the task to `.ccg/tasks/archive/2026-07/`, stage only intended files, and commit the implementation plus required task archive.
