# Project Plan: Embedded UI, Profile Storage, and Editor Workflows

## Goal
Move the UI experience into the game runtime path, remove the visible console-driven startup behavior, and replace the single profile file with a profile-per-file model under `config/` beside the patched PAIcom executable. The UI must support import, export, manage, and editor workflows without introducing a second launcher process.

## Core Decisions To Lock In
- `config/` is the only profile folder used by the runtime, and it lives next to the patched PAIcom executable.
- Each valid `.json` file in `config/` is a profile.
- Malformed JSON is still visible in Manage, but it is excluded from normal profile selection.
- Profile switching must have one clear source of truth so the active profile survives relaunch.
- Deleting a profile must only touch that profile file, and the UI must fall back to the default profile if the deleted profile was active.

## Workstream 1: Embed The UI In The Runtime
- Update [hotswap-PAIcom/Core/HotSwapTemplate.cs](hotswap-PAIcom/Core/HotSwapTemplate.cs) so the UI is hosted directly in the injected runtime path instead of spawning `PAIcomPatcher.HotSwap.Win.exe`.
- Remove the external process launch path from the runtime UI startup flow in [hotswap-PAIcom/Core/HotSwapTemplate.cs](hotswap-PAIcom/Core/HotSwapTemplate.cs).
- Review [PAIcomPatcher.HotSwap.Win.csproj](PAIcomPatcher.HotSwap.Win.csproj) for any console-window startup behavior that still belongs to the old external launcher path.
- Keep the existing command dispatch flow intact until the embedded UI path is verified end to end.
- Preserve the current injection and command dispatch structure in [hotswap-PAIcom/Core/HotSwapTemplate.cs](hotswap-PAIcom/Core/HotSwapTemplate.cs) so the runtime change does not regress command handling.

## Workstream 2: Profile Storage And Config Location
- Rework [UI-Layout/Configuration/UIProfileManager.cs](UI-Layout/Configuration/UIProfileManager.cs) so it reads and writes profiles from `config/` rather than a single fixed JSON file.
- Update [UI-Layout/Configuration/UIProfile.cs](UI-Layout/Configuration/UIProfile.cs) so the profile model cleanly represents profile metadata, button data, tab data, and the active-selection behavior needed for profile switching.
- Define a clear naming rule for profile files so import/export/manage all agree on the same profile name mapping.
- Keep the default profile fallback behavior inside [UI-Layout/Configuration/UIProfileManager.cs](UI-Layout/Configuration/UIProfileManager.cs) so broken or deleted profiles do not break the launcher.
- Document the profile folder, manual edit path, and profile-switching behavior in [README.md](README.md) and [QUICKSTART_WINFORMS_UI.md](QUICKSTART_WINFORMS_UI.md).

## Workstream 3: Files Menu Profile Actions
- Extend the File menu in [UI-Layout/CommandLauncher.cs](UI-Layout/CommandLauncher.cs) with Import, Export, and Manage entries.
- Implement Import in [UI-Layout/CommandLauncher.cs](UI-Layout/CommandLauncher.cs) as a two-step flow: choose a JSON file, then prompt for the profile name before saving into `config/`.
- Implement overwrite behavior in [UI-Layout/Configuration/UIProfileManager.cs](UI-Layout/Configuration/UIProfileManager.cs) so an exact profile-name match replaces the existing file.
- Implement Export in [UI-Layout/CommandLauncher.cs](UI-Layout/CommandLauncher.cs) as a selectable list of parsed profiles from `config/`, sorted alphabetically by the profile name starting letter.
- Implement Manage in [UI-Layout/CommandLauncher.cs](UI-Layout/CommandLauncher.cs) as a separate window that lists every JSON file in `config/`, shows parse warnings, opens the file in the Windows default editor from the pencil action, and sends deletions to the recycle bin after confirmation.
- Keep the warning indicator and tooltip logic in the Manage view tied to parse results from [UI-Layout/Configuration/UIProfileManager.cs](UI-Layout/Configuration/UIProfileManager.cs), not hard-coded UI state.

## Workstream 4: Buttons Editor Mode
- Expand the button editor path in [UI-Layout/CommandLauncherEditDialog.cs](UI-Layout/CommandLauncherEditDialog.cs) so it can edit the full button configuration requested for the UI.
- Update [UI-Layout/Commands/ButtonRegistry.cs](UI-Layout/Commands/ButtonRegistry.cs) so button edits persist through profile save/load instead of being treated as temporary view state.
- Update [UI-Layout/Layout/LayoutManager.cs](UI-Layout/Layout/LayoutManager.cs) so the displayed button order, margins, sizing, and grouping stay consistent with the saved configuration.
- Keep the edit-mode state transitions in [UI-Layout/CommandLauncher.cs](UI-Layout/CommandLauncher.cs) explicit: entering Buttons mode disables normal button activation, and leaving it restores normal clicking.
- Make discard actions revert only the unsaved changes in the edit surface, not the persisted profile file.

## Workstream 5: Layout Editor Mode
- Extend the tab and layout model in [UI-Layout/Configuration/UIProfile.cs](UI-Layout/Configuration/UIProfile.cs) so organization-tab settings can be edited independently from button settings.
- Update [UI-Layout/Components/UITabControl.cs](UI-Layout/Components/UITabControl.cs) and [UI-Layout/Layout/LayoutManager.cs](UI-Layout/Layout/LayoutManager.cs) so tab-level editing, drag-and-drop ordering, and resize behavior stay aligned with the saved profile model.
- Keep layout editing separate from button editing so the plan stays testable and the UI can validate each edit surface independently.
- Add the requested add-tab and delete-tab affordances only after the profile model can persist tab order and tab styling cleanly.

## Workstream 6: Tests And Verification
- Keep [PAIcomPatcher.Tests/PAIcomPatcher.Tests.csproj](PAIcomPatcher.Tests/PAIcomPatcher.Tests.csproj) as the dedicated Windows Forms smoke-test project.
- Expand [PAIcomPatcher.Tests/Program.cs](PAIcomPatcher.Tests/Program.cs) so it validates profile load/save, overwrite-on-import behavior, malformed JSON handling, default-profile fallback, and launcher rendering.
- Add focused assertions for [UI-Layout/Configuration/UIProfileManager.cs](UI-Layout/Configuration/UIProfileManager.cs) around round-trip serialization, command parsing, and active-profile switching.
- Keep UI smoke coverage in the test project rather than scattering ad hoc checks across the main launcher code.
- If the smoke harness grows too large, split it into smaller test files inside [PAIcomPatcher.Tests](PAIcomPatcher.Tests) instead of adding more logic to a single monolithic entrypoint.

## Documentation Work
- Update [README.md](README.md) with the new `config/` location, profile switching behavior, and manual edit workflow.
- Update [QUICKSTART_WINFORMS_UI.md](QUICKSTART_WINFORMS_UI.md) so it no longer implies the old external-launcher workflow is the production path.
- If the build or runtime launch story changes, reflect that in [SEPARATE_BUILDS.md](SEPARATE_BUILDS.md) so the hot swap and compat descriptions stay accurate.

## Acceptance Criteria
- The UI launches without spawning an external launcher process.
- No visible startup console is part of the normal UI path.
- Profiles are stored as separate JSON files under `config/` beside the patched PAIcom executable.
- Import, export, manage, and editor actions all operate on the same profile model.
- Tests cover the profile manager and the launcher smoke path in the dedicated test project.
- The documentation explains where the config files live and how to edit them safely.

## Open Questions
- Decide where the active-profile pointer should live if profile switching needs to persist independently of the selected JSON file.
- Decide whether malformed profiles should be hidden from the main launcher while remaining visible in Manage, or only flagged there.
- Confirm whether the embedded UI should remain compatible with the current command-input fallback path during migration.