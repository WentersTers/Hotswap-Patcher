# WinForms UI Quick Start Guide

## Test the Preview Form

The easiest way to verify everything works is to run the preview form:

```bash
cd "PAIcom Patch Injector"
dotnet build PAIcomPatcher.HotSwap.Win.csproj -c Release

# Launch preview
.\bin\Release\net8.0\PAIcomPatcher.HotSwap.Win.exe --preview
```

**Expected result:**
- A window titled "Command Launcher - PREVIEW MODE" appears
- Buttons arranged in a grid showing shortened versions of command phrases
- Button text is shortened from the commands.txt file
- Each button shows preview text when clicked
- Status bar updates with dispatch information

---

## Integrate with Patched Runtime

To make the UI available within the patched PAIcom.exe:

### 1. Create a Launcher Batch/Script
```batch
@echo off
REM launcher.bat - placed in game directory
cd /d "%~dp0"
start PAIcomPatcher.HotSwap.Win.exe --launcher "%cd%"
```

Or in PowerShell:
```powershell
# launcher.ps1
$gameDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$gameDir\PAIcomPatcher.HotSwap.Win.exe" --launcher $gameDir
```

### 2. Call from HotSwapTemplate (Optional)
The HotSwapTemplate can be extended to show the UI. Add to the template:
```csharp
// In HotSwapRuntime - after initialization
System.Diagnostics.Process.Start("launcher.bat");
// or
System.Diagnostics.Process.Start("powershell.exe", "-File launcher.ps1");
```

---

## File Structure for Deployment

```
PAIcom.exe (patched with hotswap)
command_input.txt
custom-commands/
  commands.txt
  audio/
animations/
PAIcomPatcher.HotSwap.Win.exe  ← Command launcher executable
launcher.bat  ← Optional: starts the UI
hotswap.log  ← Logs from runtime dispatch
```

---

## Testing Workflow

### Step 1: Build
```bash
dotnet build PAIcomPatcher.HotSwap.Win.csproj -c Release
```

### Step 2: Test Preview UI
```bash
.\bin\Release\net8.0\PAIcomPatcher.HotSwap.Win.exe --preview
```

### Step 3: Verify Commands Load
- Check that buttons appear for each command in commands.txt
- Very long command phrases should be shortened with "…"
- No commands should show "No commands found" message

### Step 4: Test Button Click
- Click a button
- In preview mode, you'll see a test notification
- In production mode, the command would be dispatched

### Step 5: Check Files
- After clicking, `command_input.txt` should contain the command phrase
- Check `hotswap.log` for dispatch entries (in the game directory)

---

## Troubleshooting

### Buttons don't appear
- Check that `custom-commands/commands.txt` exists and is readable
- Verify the file format: `phrase (file.txt)` on each line
- Comments starting with `#` are ignored

### Commands not dispatching
- Verify command_input.txt is being written (it updates after each click)
- Check hotswap.log for error messages
- Ensure the target directory structure matches:
  ```
  BaseDir/
    custom-commands/commands.txt
    command_input.txt
  ```

### UI doesn't start
- Ensure System.Windows.Forms and System.Drawing.Common packages installed
- Run preview form first to isolate issues
- Check that .NET 8.0 runtime is available

---

## Advanced Usage

### Custom Base Directory
```bash
PAIcomPatcher.HotSwap.Win.exe --launcher "D:\Custom\Game\Path"
```

### From C# Runtime Code
```csharp
// If you need to show the UI from within PAIcom
var launcher = new CommandLauncher();
launcher.ShowDialog();

// Or refresh commands if commands.txt was edited
launcher.RefreshCommands();
```

### Direct Dispatch from Code
```csharp
// If HotSwapRuntime is accessible
HotSwapRuntime.DispatchCommandPhrase("hey paicom open youtube");
```

---

## Platform Support

- **Windows**: Full support (.NET 8.0)
- **Linux/Mac**: Limited (WinForms is Windows-only, but file dispatch works)
  - Use `--launcher` for command dispatch without UI

## Files to Review

- [CommandLauncher.cs](UI/CommandLauncher.cs) - Main UI component
- [CommandLauncherPreview.cs](UI/CommandLauncherPreview.cs) - Test form
- [HotSwapTemplate.cs](Core/HotSwapTemplate.cs) - Runtime dispatch logic
- [Program.cs](Program.cs) - Entry point and UI launcher
