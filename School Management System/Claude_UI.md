# Claude.md — School Management System (WPF VB.NET) Agent Rules

## Project Context
- Tech: WPF Desktop App, VB.NET (.NET 10).
- Current Views that already exist:
  - `Views/LoginWindow.xaml`
  - `Views/AdminDashboardWindow.xaml`
- Styling/Theming already exists in `Resources/Styles` via ResourceDictionaries.
- We are focusing on the **Admin** side UI first.
- Reference spec file in repo: `AdminPagesSpec.md` (source of truth for page contents).

## Current Admin Modules (Pages)
- Dashboard — overall overview
- Courses — create/manage courses
- Subjects — create/manage subjects (belongs to Course)
- Sections — create/manage sections
- Rooms — create/manage rooms
- Students — manage students (enroll per subject later)
- Teachers/Professors — manage teachers, filter by courses, days, sections, rooms
- Scheduling — manages class schedules with tabs (Class Scheduling / Teacher View / Section View)

## .NET 10 Compatibility (Critical)
This project targets **`net10.0-windows`** (.NET 10, not .NET Framework). Assembly references that work on .NET Framework will **fail at runtime** on .NET (Core). Follow these rules strictly:

- **NEVER** use `assembly=mscorlib` — it does not exist in .NET 10.
- **NEVER** use `assembly=PresentationCore` or `assembly=PresentationFramework` in `clr-namespace` XAML references for BCL types.
- For `System` namespace types in XAML (e.g., `sys:String`, `sys:Boolean`, `sys:Double`), use:
  ```xml
  xmlns:sys="clr-namespace:System;assembly=System.Runtime"
  ```
- For `System.Collections` types, use `assembly=System.Collections`.
- For `System.ComponentModel` types (e.g., `INotifyPropertyChanged`), import in VB code — do **not** reference `System.dll` or `System.ObjectModel.dll` in XAML.
- When in doubt, check the assembly name via `dotnet` docs or use the `System.Runtime` assembly as the default for core BCL types.

## Absolute Rules (Non-Negotiable)
1) **FRONTEND ONLY**
   - No database code (MySQL), no repositories, no services, no API calls.
   - No authentication logic. Login remains UI-only for now.
2) **ONE PAGE PER REQUEST**
   - Implement only the requested page/module per prompt.
   - Do not build multiple pages unless explicitly asked.
3) **NO MOCK DATA**
   - Do NOT add fake/demo/sample records to any list.
   - Collections must start empty.
   - Use clean **empty states** in the UI (e.g., “No records yet”).
4) **SAME COLORS / THEME**
   - Do **not** introduce new colors.
   - Do **not** hardcode colors/brushes (`#hex`, `rgb`, etc.) in new pages.
   - Use existing `StaticResource` / `DynamicResource` brushes and existing styles from `Resources/Styles`.
   - Use existing fonts, corner radii, shadows, spacing patterns already used in the app.
5) **Minimal Code-Behind**
   - Code-behind is allowed only for UI behaviors (window drag, toggles, simple visual states, file dialog preview).
   - No business logic, no data generation, no persistence.
6) **Navigation**
   - New pages must display inside `AdminDashboardWindow` in a content host.
   - If `AdminDashboardWindow` does not already have a `ContentControl` or `Frame`, add ONE named `MainContentHost` and keep the existing dashboard layout intact.
   - Sidebar buttons must navigate to the page (no functional backend required).
7) **No Unnecessary Refactors**
   - Do not restructure unrelated files.
   - Do not rename existing styles/resources.
   - Do not add new NuGet packages unless explicitly requested.

## Folder & File Conventions
- Pages go in: `Views/Pages/`
  - Example: `Views/Pages/StudentsPage.xaml`
- ViewModels go in: `ViewModels/`
  - Example: `ViewModels/StudentsViewModel.vb`
- Models (only if needed) go in: `Models/`
  - Example: `Models/StudentListItem.vb`
- Keep naming consistent and readable.

## MVVM Guidance (Lightweight, Future-Ready)
- Each page should have:
  - A ViewModel with:
    - Empty `ObservableCollection(Of T)` for list views/grids (do not populate)
    - Initialize collections in the constructor (empty)
    - `SelectedItem` property
    - Filter properties where needed (SearchText, SelectedCourse, etc.)
    - `ICommand` properties (can be stub commands that do nothing yet)
- Avoid heavy frameworks; simple VB.NET MVVM patterns are fine.

## Empty State Requirements (Important)
If a page includes a DataGrid/List:
- Provide a dedicated empty-state panel:
  - Title text like “No records yet”
  - Supporting text like “Add a new item to get started.”
  - Optional “New” button (UI-only)
- Empty state must appear when the list has 0 items.

### Required Implementation Pattern (NO converters)
Use an overlay empty state triggered by the control’s `Items.Count`:

- Wrap the DataGrid and the empty-state panel in a parent `Grid`.
- Name the DataGrid (example: `x:Name="RecordsGrid"`).
- Overlay a `Border` (empty state) on top of the DataGrid.
- Toggle the overlay using a `DataTrigger` bound to:
  `ElementName=RecordsGrid, Path=Items.Count`

Example (adjust brush/style keys to existing theme resources; do NOT hardcode colors):
```xml
<Grid>
  <DataGrid x:Name="RecordsGrid"
            ItemsSource="{Binding Records}"
            AutoGenerateColumns="False"
            IsReadOnly="True" />

  <Border Padding="24"
          CornerRadius="12"
          Background="{DynamicResource CardBackgroundBrush}">
    <Border.Style>
      <Style TargetType="Border">
        <Setter Property="Visibility" Value="Collapsed"/>
        <Style.Triggers>
          <DataTrigger Binding="{Binding ElementName=RecordsGrid, Path=Items.Count}" Value="0">
            <Setter Property="Visibility" Value="Visible"/>
          </DataTrigger>
        </Style.Triggers>
      </Style>
    </Border.Style>

    <StackPanel HorizontalAlignment="Center" Width="340">
      <TextBlock Text="No records yet"
                 FontSize="18"
                 FontWeight="SemiBold"
                 TextAlignment="Center"/>
      <TextBlock Margin="0,8,0,0"
                 Text="Connect the database or add your first record to see results here."
                 Opacity="0.8"
                 TextWrapping="Wrap"
                 TextAlignment="Center"/>
      <Button Margin="0,16,0,0"
              Content="New"
              Command="{Binding NewCommand}"
              HorizontalAlignment="Center"/>
    </StackPanel>
  </Border>
</Grid>
```

- If `CardBackgroundBrush` does not exist, reuse an existing brush key already used in the app. Do NOT create new brush resources.

## Output Format Requirement
When you respond with code:
- Group output by file path using clear headings, e.g.:

  `### Views/Pages/StudentsPage.xaml`
  ```xml
  ...
  ```

  `### ViewModels/StudentsViewModel.vb`
  ```vb
  ...
  ```

- Include only files that were added/changed.
- Ensure the project builds with no missing references.

## Page-by-Page Implementation Checklist
For each requested page:
- [ ] Create the Page XAML using existing styles and brushes (no new colors)
- [ ] Create matching ViewModel (empty collections initialized, selected item, commands)
- [ ] Add minimal navigation hookup from `AdminDashboardWindow`
- [ ] Add empty-state UI for all lists/grids using the required overlay pattern
- [ ] Confirm no mock data exists anywhere
- [ ] Confirm compile-safe code

## Example Navigation Expectations
- Sidebar button “Students” loads `StudentsPage` into `MainContentHost`.
- Default load behavior should not be changed unless explicitly asked.

---

## Reminder
Only implement what the prompt asks. If details are missing, choose the safest minimal UI approach:
- Keep lists empty
- Show empty state using the required overlay pattern
- Provide form fields and buttons visually (wired to commands but no backend)
- Preserve the existing theme and structure
