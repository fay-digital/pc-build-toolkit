# pc-build-toolkit

A PowerShell + WPF deployer for PC build validation. Runs on a fresh Windows install with no setup — fetched and executed from a single one-liner. Installs a curated set of diagnostic and benchmarking tools, applies common bench tweaks, and produces a per-build report.

Built for Fay Digital's PC building workflow. Public so it can be run from any bench machine without auth.

---

## Quick start

From an **elevated PowerShell** prompt on the target machine:

```powershell
irm fay.digital/pbt | iex
```

The script will self-elevate if you forget — a UAC prompt appears, approve it, and the UI launches.

That's the whole installation. Nothing persists on disk (apart from the log file and the bench report the tool itself writes).

---

## What it does

The UI has two panels: **applications** (install selectors) and **system tweaks** (opt-in toggles). Pick what you want, click **Run**, watch the log and progress bar.

### Application catalog

| App | Source | Default |
|---|---|---|
| AIDA64 Extreme | winget | ✓ |
| HWiNFO | winget | ✓ |
| CrystalDiskMark | winget | ✓ |
| Cinebench R23 | winget | ✓ |
| 3DMark (Steel Nomad) | zip bundle | ✓ |
| FurMark 2 | winget | |
| OCCT | winget | |
| CPU-Z | winget | |
| GPU-Z | winget | |
| Prime95 | winget | |

### System tweaks

- Set power plan timeouts (display, sleep, hibernate) to **never**
- Disable hibernation (`powercfg -h off`)
- Clear the Downloads folder
- Empty the Recycle Bin
- Clear browser history (Edge, Chrome, Firefox)

### Other actions

- **Uninstall all** — removes every catalog app currently present on the machine. Apps not installed are skipped gracefully.
- **Bench report** — writes `SoftwareReport_<HOSTNAME>_<TIMESTAMP>.txt` to the Desktop after every run. Includes CPU, GPU, RAM, motherboard, storage, and a tick-list of every action taken. On by default; togglable.

---

## Requirements

- **Windows 10 (1809 or later) or Windows 11**
- **PowerShell 5.1+** (ships with Windows — no install needed)
- **Admin rights** (the script self-elevates)
- **Internet connection**
- **~10 GB free on C:** (required by pre-flight check; needed for the 3DMark bundle download + extraction + install)

No dependencies to install upfront. `winget` ships with supported Windows versions. Chocolatey is bootstrapped automatically if and only if a selected app needs it.

---

## How it works

**Multi-source installer.** Each catalog entry declares a source type:

- `winget` — standard `winget install --id <id>` with silent flags
- `choco` — `choco install <id> -y`, auto-bootstraps Chocolatey if missing
- `direct` — downloads an installer EXE from a URL, runs it silently
- `zip` — downloads a zip, extracts, runs the named setup executable from inside. Used for 3DMark because the installer needs its DLC files as siblings at install time.

**Uninstall.** For `winget` and `choco` sources, uses the respective uninstall command. For `direct` and `zip`, searches the registry's `Uninstall` keys for a `DisplayName` matching the configured pattern and runs whatever `UninstallString` Windows recorded, with `/S` appended for silent mode.

**Pre-flight checks.** Before running, verifies internet reachability, disk space, and winget health. Fails loudly with a log message rather than silently proceeding into a broken install.

**Background execution.** The install pipeline runs on a separate PowerShell runspace so the UI stays responsive during multi-minute operations. Log lines appear as they happen; progress bar ticks per completed task.

**Self-update check.** On launch, compares the running script's SHA-256 against the raw URL and logs a notice if they differ. Doesn't auto-update — just tells you there's something newer.

**Logging.** Everything goes to `%TEMP%\software.log` in addition to the on-screen log. Survives the script exiting.

---

## Forking and customising

The catalog is the top of the script (around line 32). Each entry is a hashtable:

```powershell
@{ Id='Vendor.App'; Name='App Name'; Category='Benchmark'; Source='winget' }
```

Swap in your own apps, change the defaults, reorder. `Source='zip'` and `Source='direct'` entries take extra fields — see the 3DMark entry as a template.

To point the one-liner at your own fork, change the URL in your launch command. Nothing else is hard-coded to the `fay-digital` namespace except comments.

---

## Known limitations

- **Cinebench R23** sometimes fails to launch after install due to a missing Intel C++ runtime DLL. Upstream winget packaging issue. Workaround: run Cinebench from its actual install folder (`%LOCALAPPDATA%\Microsoft\WinGet\Packages\Maxon.CinebenchR23_*\`) rather than the Start Menu shortcut.
- **3DMark bundle only includes Steel Nomad.** Adding Fire Strike or other DLCs would push the zip over GitHub's 2 GB release-asset limit. Self-host if you need more tests.
- **Winget IDs drift.** If a catalog entry fails with exit code `-1978335212` ("id not found"), the winget manifest has moved. Run `winget search <app name>` to find the current id and update the catalog. OCCT is the most likely candidate for this.
- **Browser history clear fails if the browser is running** — files are locked. The script logs a warning and moves on. Close browsers first if it matters.

---

## Repository layout

```
.
├── deployer.ps1        # the script — this is what gets executed
├── README.md           # this file
└── (release assets)    # 3dmark-bundle.zip attached to release tag v1.0.0
```

Releases hold the 3DMark bundle as an attached asset. The script's catalog references it by URL.

---

## Versioning

Standard [semver](https://semver.org). Patch for fixes, minor for new features, major for breaking changes. See the [Releases](https://github.com/fay-digital/pc-test-suite/releases) page for per-version changelogs.

Pin a specific version by swapping `main` in the one-liner for a release tag:

```powershell
# always latest
irm https://raw.githubusercontent.com/fay-digital/pc-test-suite/main/deployer.ps1 | iex

# frozen at v1.0.0
irm https://raw.githubusercontent.com/fay-digital/pc-test-suite/v1.0.0/deployer.ps1 | iex
```

---

## Safety notes

This script runs as **Administrator** and does things that are hard to undo: disabling hibernation, clearing Downloads, emptying the Recycle Bin, uninstalling software. Read the options before clicking Run. The **Uninstall all** button asks for confirmation but is destructive — it will uninstall every catalog app currently on the machine.

The script is fetched from GitHub over HTTPS every time you run the one-liner, so whoever controls this repo controls what runs on your bench machines. If you're forking this for your own use, host from a repo you control.

---

## License

MIT. See [LICENSE](LICENSE).
