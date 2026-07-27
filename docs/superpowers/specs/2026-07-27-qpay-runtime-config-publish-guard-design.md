# QPay Runtime Config Publish Guard Design

## Background

`QPayBackend` is an ASP.NET Core 2.1 application targeting .NET Framework 4.7.1. IIS starts `QPayBackend.exe`, while `ToolUtility` reads settings through `ConfigurationManager.AppSettings`. The effective production configuration is therefore `QPayBackend.exe.config` beside the executable.

The existing publish output contains both source `app.config` and generated `QPayBackend.exe.config`. Updating only `app.config` looks valid to an operator but has no effect on the running process. This caused the current incident: the source file contained updated `611_XKeyID` and `NANKAN_XKeyID` values while the runtime file still contained the previous values.

## Goal

Make it impossible to produce an apparently successful QPay deployment when the generated runtime `appSettings` differ from the source settings, and remove the misleading source `app.config` from deployment packages.

## Non-goals

- Do not change QPay payment, CRM, or LINE processing behavior.
- Do not modify MyPay.
- Do not migrate secrets to a new secret store in this change.
- Do not log or print any setting value, complete XML document, XKey, Token, password, or payment data.
- Do not add a runtime fail-fast guard in this first change; prevention belongs at the publish boundary where source and generated configurations are both available.

## Approaches Considered

### Documentation only

Documenting that `QPayBackend.exe.config` is effective is inexpensive, but it still relies on memory and cannot stop an invalid deployment. This is insufficient.

### Publish-time hard guard — selected

Keep `app.config` as the source of truth, ensure the Web SDK generates `QPayBackend.exe.config`, compare their `appSettings` after publish, and fail publication on any discrepancy. Remove `app.config` from the deployment output after successful validation.

This is the smallest change that blocks the exact failure mode without modifying production payment logic.

### Configuration-system migration

Moving XKeys to environment variables or a secret vault would remove file ambiguity, but it requires broader runtime and deployment changes. It should be evaluated separately.

## Design

### Project configuration

`QPayBackend/QPayBackend.csproj` will explicitly designate `QPayBackend/app.config` as the application configuration source and mark the source file as never copied to output or publish directories. The SDK must continue generating `QPayBackend.exe.config` from it.

An `AfterTargets="Publish"` target will invoke a repository-owned PowerShell verifier. Because the target is part of the project, direct `dotnet publish` calls cannot bypass validation.

### Configuration verifier

`QPayBackend/Verify-PublishedConfig.ps1` will accept three explicit literal paths:

- source `app.config`;
- generated `QPayBackend.exe.config`;
- forbidden published `app.config`.

It will parse only `/configuration/appSettings/add` and compare dictionaries using ordinal, case-sensitive, whitespace-sensitive string comparison. It will fail when:

- either required file is missing;
- XML cannot be parsed;
- a key is missing or unexpectedly added;
- a value differs;
- a required `SINOPAC_SITE`, `SANDBOX_SITE`, or `*_XKeyID` value is empty or contains outer whitespace;
- source `app.config` appears in the deployment output.

Errors will contain only the affected key name and a fixed error category. Values and raw exception details will never be emitted. Any failure exits nonzero so MSBuild publication fails.

### Clean release entry point

`QPayBackend/DotNetPublish-Release.bat` will become a thin wrapper that:

1. resolves paths relative to the batch file;
2. runs `dotnet clean -c Release`;
3. publishes to a newly created, timestamped directory under the ignored `artifacts` directory;
4. preserves and returns the failing command's exit code;
5. never treats `pause` as publication success.

The generated deployment directory must contain `QPayBackend.exe.config` and must not contain `app.config`.

### Test strategy

A dependency-free PowerShell test harness will create temporary XML fixtures and execute the real verifier. Tests will cover:

- equal source/runtime `appSettings` succeed;
- changed value fails while output contains only the key name, not either value;
- missing key fails;
- outer whitespace in XKey fails;
- missing runtime configuration fails;
- forbidden published `app.config` fails.

An integration publish test will publish to a fresh temporary directory and assert that `QPayBackend.exe.config` exists, `app.config` is absent, and project publication exits successfully. The full solution will then be built with Visual Studio MSBuild.

## Error handling and safety

- All filesystem paths are explicit and resolved with literal-path APIs.
- Temporary test and publish directories are created beneath known test/artifact roots.
- No recursive deletion targets the repository root, user profile, drive root, or an unresolved path.
- Publish validation stops before a deployable package is accepted.
- Existing uncommitted or committed XKey values are treated as opaque secrets and are never included in test fixtures, logs, documentation, or review prompts.

## Acceptance criteria

- A normal Release publish generates `QPayBackend.exe.config` from source settings.
- The deployment output contains no `app.config`.
- Any mismatch or unsafe whitespace causes a nonzero publish result.
- Diagnostic output identifies only the key and failure category.
- Tests demonstrate the failing cases before implementation and pass afterward.
- QPay, CRM, LINE, and MyPay production code remains unchanged.
