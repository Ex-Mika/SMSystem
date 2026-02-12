# AGENTS.md

## Project
- Name: School Management System
- Stack: VB.NET + WPF
- Current priority: Frontend-first implementation (login UI and visual flows)

## Working Rules
- Build UI features first before backend integration.
- Do not add mock database layers or fake backend services unless explicitly requested.
- Keep changes modular and maintainable.
- Preserve the current visual language (blue gradient theme, card-based login, PRMSU branding).

## Current Structure
- `School Management System/Views/` -> windows and view markup/code-behind
- `School Management System/Resources/Styles/` -> shared colors and control styles
- `School Management System/Resources/Images/` -> logo, background, and icon assets

## UI Conventions
- Reuse shared styles in `Resources/Styles` instead of duplicating inline templates.
- Keep icons visually consistent (size, color, alignment).
- Validate title bar controls (minimize, maximize/restore, close) after UI changes.
- Keep layout responsive for common desktop sizes.
- Do not use 4k images only maximum of 1920x1080

## Code-Behind Conventions
- Keep code-behind focused on UI behavior only.
- Avoid business logic, data access, and service calls in views during frontend phase.
- Use small helper methods for control state updates (visibility, labels, icon state).
- Guard event handlers against early initialization null references.

## Verification
- After UI/code changes, run:
  - `dotnet build "School Management System.slnx"`
- Ensure:
  - No compile errors
  - Role-based label text updates work
  - Password show/hide toggle works
  - Custom title bar behavior works (drag, double-click maximize/restore, control buttons)

## Asset Notes
- Prefer project-local assets under `Resources/Images`.
- Register new image assets in `School Management System.vbproj` as `Resource`.
- Use appropriately sized assets for UI usage to avoid unnecessary file bloat.

