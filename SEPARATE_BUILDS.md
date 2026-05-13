# Separate Build Structure

The PAIcom Patcher project is now split into two fully independent builds for easier sharing and distribution.

## Folder Layout

```
PAIcom Patch Injector/
├── hotswap-PAIcom/          ← HotSwap build (Windows-only with file-watching)
│   ├── Core/                ← All .cs files including HotSwap-specific
│   ├── animations/          ← Animation scripts
│   ├── custom-commands/     ← Custom command mappings
│   ├── audio/               ← Audio files
│   ├── files/               ← Data files
│   ├── UI/                  ← UI resources
│   ├── Program.cs           ← Entry point (HotSwap variant)
│   ├── PAIcomPatcher.HotSwap.Win.csproj
│   ├── BUILD.md             ← Build instructions for HotSwap
│   └── README.md
│
├── compat-PAIcom/           ← Compat build (Cross-platform, no hotswap)
│   ├── Core/                ← Shared .cs files (HotSwap files excluded)
│   ├── animations/          ← Animation scripts
│   ├── custom-commands/     ← Custom command mappings
│   ├── audio/               ← Audio files
│   ├── files/               ← Data files
│   ├── UI/                  ← UI resources
│   ├── Program.cs           ← Entry point (Compat variant)
│   ├── PAIcomPatcher.Compat.csproj
│   ├── BUILD.md             ← Build instructions for Compat
│   └── README.md
│
└── (original root files)    ← Original structure remains for reference
```

## Key Differences

| Aspect | HotSwap | Compat |
|--------|---------|--------|
| **Location** | `hotswap-PAIcom/` | `compat-PAIcom/` |
| **Platforms** | Windows only (`win-x64`) | All (`win-x64`, `linux-x64`, `osx-x64`, `osx-arm64`) |
| **Command-input** | ✅ Yes (`command_input.txt`) | ❌ No |
| **Hotswap Runtime** | ✅ Yes | ❌ No |
| **Core Files** | 9 files (includes HotSwap logic) | 6 files (basic patching) |
| **Output Binary** | `PAIcomPatcher.HotSwap.Win.exe` | `PAIcomPatcher.Compat` (per platform) |

## Contents of Each Build

### Shared Between Both
- `AssemblyAnalyzer.cs` - IL analysis  
- `AssemblyPatcher.cs` - Patch orchestration
- `ILPatternScanner.cs` - IL pattern matching
- `LauncherGenerator.cs` - OS launcher scripts
- `ReferenceAssemblyResolver.cs` - Assembly resolution
- `VoskResourceEmbedder.cs` - Resource embedding
- `Core/ManagedLibraries/` - .NET Framework refs
- `Core/NativeLibraries/` - Native dependencies
- `animations/`, `custom-commands/`, `files/`, `audio/`, `UI/` - Data

### HotSwap Only
- `HotSwapInjector.cs` - Runtime injection logic
- `HotSwapTemplate.cs` - Generated runtime template
- `CommandHandlerFinder.cs` - Command handler IL detection
- All conditional `#if HOTSWAP_BUILD` code paths

### Compat Only
- None (this build excludes hotswap logic)
- All conditional `#if COMPAT_BUILD` code paths

## Building Each Variant

### HotSwap (Windows-only with file-watching)
```bash
cd hotswap-PAIcom
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
# Output: bin/Release/net8.0/win-x64/publish/PAIcomPatcher.HotSwap.Win.exe
```

### Compat (Cross-platform, basic patching)
```bash
cd compat-PAIcom
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
# Output: bin/Release/net8.0/win-x64/publish/PAIcomPatcher.Compat
```

For other platforms, replace `-r win-x64` with:
- `linux-x64` (Linux)
- `osx-x64` (macOS Intel)
- `osx-arm64` (macOS Apple Silicon)

## Sharing Individual Builds

### To Share Only HotSwap
```bash
cd hotswap-PAIcom
# Zip and distribute
7z a -r PAIcomPatcher.HotSwap.Win.7z *
```

### To Share Only Compat
```bash
cd compat-PAIcom
# Zip and distribute
7z a -r PAIcomPatcher.Compat.7z *
```

## Getting Started

1. Choose which build you need:
   - **HotSwap** for file-watching command input (Windows)
   - **Compat** for cross-platform assembly analysis

2. Navigate to that folder:
   ```bash
   cd hotswap-PAIcom    # or: cd compat-PAIcom
   ```

3. Read `BUILD.md` for detailed build instructions

4. Build and run as described in `BUILD.md`

## Compilation-Time Separation

Both builds use **conditional compilation** via `DefineConstants` in their `.csproj` files:

**HotSwap:**
```xml
<DefineConstants>$(DefineConstants);HOTSWAP_BUILD</DefineConstants>
```

**Compat:**
```xml
<DefineConstants>$(DefineConstants);COMPAT_BUILD</DefineConstants>
```

Code uses `#if HOTSWAP_BUILD` and `#if COMPAT_BUILD` to conditionally include/exclude logic.

## Notes

- The original root structure is preserved for reference and version control
- Each build folder is fully self-contained for independent distribution
- Shared files are physically copied (not linked) for portability
- Data folders (animations, commands, etc.) are duplicated in each build

For detailed information on each variant, see:
- [hotswap-PAIcom/BUILD.md](hotswap-PAIcom/BUILD.md)
- [compat-PAIcom/BUILD.md](compat-PAIcom/BUILD.md)
