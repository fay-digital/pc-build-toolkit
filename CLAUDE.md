# PC Build Toolkit — Claude Context

## What this is

A single-file PowerShell + WPF GUI app (`deployer.ps1`) that PC builders run on fresh Windows installs to install diagnostics/benchmarks, remove bloat, and produce a per-build hardware report. Delivered as a one-liner:

```powershell
irm fay.digital/pbt | iex
```

Everything lives in `deployer.ps1`. There is no build step, no package manager, no compilation.

## Repository layout

```
deployer.ps1          # the entire application (~2000 lines)
README.md
.github/workflows/
  ci.yml              # PSScriptAnalyzer lint + XAML validation
```

## Script structure (by line range)

| Range | Content |
|---|---|
| 1–35 | Version, URLs, self-elevation |
| 36–58 | `$script:AppCatalog` — install manifest |
| 59–79 | `$script:BloatList` — AppX removal list |
| 80–955 | XAML UI definition (inline here-string) |
| 956–1055 | `Write-Log`, `Write-UiLog`, logging helpers |
| 1056–1165 | Pre-flight, self-update, UI state helpers |
| 1166–1465 | Download, extraction, install/uninstall functions |
| 1466–1641 | Bloat removal, Start menu, Windows Update |
| 1642–1778 | System tweaks (power, hibernation, browser clear) |
| 1779–1912 | `New-BenchReport` — hardware report generation |
| 1913–end | `Start-Pipeline` — main orchestration + UI event wiring |

## Development workflow

**Never commit directly to `main`.** Always:

1. Work on a feature branch (`claude/description-of-change`)
2. Commit with clear messages
3. Push and open a PR
4. CI must pass (PSScriptAnalyzer + XAML) before merging

Branch naming: `claude/<short-description>`

## CI

GitHub Actions runs on every push and PR:

- **PSScriptAnalyzer** — `Error` and `Warning` severity, with these rules excluded:
  `PSAvoidUsingWriteHost`, `PSUseShouldProcessForStateChangingFunctions`,
  `PSAvoidUsingComputerNameHardcoded`, `PSAvoidOverwritingBuiltInCmdlets`, `PSUseSingularNouns`
- **XAML validation** — parses both standalone `.xaml` files and embedded XAML here-strings in `.ps1` files

CI runs on `windows-latest`. There is no local test runner — CI is the test gate.

## Code conventions

- PowerShell 5.1 compatible (no `#Requires -Version 7`, no `??=`, no ternaries)
- `$script:` scope for module-level state shared across functions
- UI updates always go through `Dispatcher.Invoke` (script runs UI on main thread, pipeline on runspace)
- Log to both `Write-Log` (file) and `Write-UiLog` (on-screen) for anything user-visible
- Prefer `Write-UiLog` with `'OK'`, `'WARN'`, or `'ERROR'` level over raw text
- No positional parameters on non-trivial functions — always use `param()`
- `$sync` hashtable is the runspace-safe shared state between UI thread and background pipeline

## Adding an app to the catalog

Each entry is a hashtable in `$script:AppCatalog`. Minimum fields:

```powershell
@{ Id='Vendor.App'; Name='Display Name'; Category='Benchmark'; Source='winget' }
```

For `zip` source, also add: `DownloadUrl`, `SetupExecutable`, `SilentArgs`, `UninstallRegistryMatch`.

Add the `Id` to `$script:DefaultChecked` if it should be pre-ticked.

## Key things to avoid

- Do not change the XAML structure without checking `$sync` references — element names in XAML must match `$sync.<ElementName>` lookups in the script body
- Do not use `Write-Host` for functional output (PSScriptAnalyzer warning, and it doesn't reach the log file)
- Do not use `Start-Process` with `-Wait` on the UI thread — it freezes the window
- The script must remain a single file (no dot-sourcing, no modules) — it is fetched and executed directly

## Version

Current: `v1.1.0` (line 15: `$SCRIPT_VERSION`). Bump on meaningful changes.
Format: semver — patch for fixes, minor for new features, major for breaking changes.
