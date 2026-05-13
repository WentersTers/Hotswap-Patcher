# Compat Build - Build Instructions

This is the **Compat** variant with cross-platform support but no hotswap/command-input injection.

## Features
- Cross-platform (`win-x64`, `linux-x64`, `osx-x64`, `osx-arm64`)
- Basic assembly patching (no hotswap logic)
- No file-watching or command-input injection
- Lighter weight, simpler codebase
- Suitable for interoperability research

## Building

### Prerequisites
- .NET 8 SDK or later
- Command line: `dotnet --version` should return 8.0+

### Build Commands

**Development build (Debug):**
```bash
dotnet build PAIcomPatcher.Compat.csproj
```

**Release build:**
```bash
dotnet build -c Release PAIcomPatcher.Compat.csproj
```

**Platform-specific self-contained executable:**

Windows:
```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Linux:
```bash
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true
```

macOS (Intel):
```bash
dotnet publish -c Release -r osx-x64 --self-contained true -p:PublishSingleFile=true
```

macOS (Apple Silicon):
```bash
dotnet publish -c Release -r osx-arm64 --self-contained true -p:PublishSingleFile=true
```

Output: `bin/Release/net8.0/{RID}/publish/PAIcomPatcher.Compat`

## Usage

```bash
./PAIcomPatcher.Compat <path-to-PAIcom.exe> [options]

Options:
  --out <file>    Output path (default: <input>.patched.exe)
  --dry-run       Analyse and report patch points without writing output
  --backup        Copy original to <input>.bak before patching
  --verbose       Print detailed IL scan progress
  --analyze       Scan assembly without patching
```

## What This Build Does

This is a **basic assembly patcher** that:
1. Loads the target PE/IL assembly
2. Analyzes IL code patterns
3. Reports potential patch points
4. (Does NOT inject hotswap runtime)

It's designed for research and interoperability scenarios rather than runtime injection.

## No Hotswap Features

This build intentionally excludes:
- ❌ `HotSwapInjector.cs`
- ❌ `HotSwapTemplate.cs`
- ❌ `CommandHandlerFinder.cs`
- ❌ File-watching runtime
- ❌ `command_input.txt` support
- ❌ Animation scripts

## When to Use This Build

- Cross-platform compatibility needed
- Pre-built deployments for Linux/macOS (via Wine)
- Assembly analysis without injection
- Research or educational purposes
- Minimal runtime footprint

## Troubleshooting

- **Build fails with "COMPAT_BUILD not defined"** - Ensure you're using `PAIcomPatcher.Compat.csproj`
- **Unsupported platform** - Only `win-x64`, `linux-x64`, `osx-x64`, `osx-arm64` are supported
