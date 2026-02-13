# AdminPagesSpec.md — School Management System (WPF VB.NET) Frontend-Only Page Content Spec

> Use this document as the single source of truth for what each **Admin** page must contain.
> Constraints: **Frontend only**, **no mock data**, **same existing theme/colors**, **one page per request**.

---

## Global Constraints (Apply to ALL Pages)

### 1) Frontend Only
- No MySQL code, no repositories/services, no network calls.
- No authentication logic (Login remains UI-only for now).

### 2) No Mock Data
- Do not seed demo records.
- All `ObservableCollection` lists start empty.
- Every list/grid must show a clean empty-state fallback.

### 3) Same Theme / Same Colors
- Do not add new colors or hardcode hex values.
- Use existing `StaticResource`/`DynamicResource` brushes and styles from `Resources/Styles`.
- Reuse existing control styles (Buttons/TextBoxes/ComboBoxes/Cards).

### 4) Navigation & Hosting
- All new pages are hosted inside `AdminDashboardWindow` in a content host.
- If missing, add **one** host named `MainContentHost` (Frame or ContentControl) without changing existing layout.
- Sidebar buttons navigate to pages.

### 5) Minimal Code-Behind
- Only UI behaviors (visual toggles, file dialog preview) are allowed.
- No business logic, no data generation, no persistence.

---

## Standard UI Patterns

### A) Empty State Overlay (Required for Grids/Lists)
**Always implement empty states as an overlay triggered by `Items.Count` (NO converters).**

Template (replace resource keys with ones that exist in your project):
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

> If `CardBackgroundBrush` does not exist, reuse a brush already used by the app’s cards/panels. Do NOT invent new brushes.

### B) MVVM (Lightweight)
Each page should have:
- `ObservableCollection(Of T)` list property (initialized empty)
- `SelectedItem` property
- Filter properties (SearchText, SelectedCourse, etc.) where applicable
- ICommand stubs (New/Edit/Save/Cancel/Clear/Refresh)

---

# Final Admin Page List (Agreed)

1. Dashboard
2. Courses
3. Subjects
4. Sections
5. Rooms
6. Students
7. Teachers/Professors
8. Scheduling (Tabs)

> Notes:
> - **Subjects belong to a Course** (YES)
> - **Sections are real entities** (YES)
> - Students will **enroll per subject later** (so no “assign student to section” as the core model)
> - Scheduling uses **Tabs** (YES)

---

## 1) Dashboard — Overview

### Purpose
Quick snapshot + shortcuts. Values can be placeholders (e.g., “—”) until backend exists.

### Content
- KPI Cards (4–6):
  - Total Students
  - Total Teachers
  - Total Courses
  - Total Subjects
  - Total Sections
  - Scheduled Classes Today (optional)
- Quick Actions:
  - Add Student
  - Add Teacher
  - Create Course
  - Create Subject
  - Open Scheduling
- Panels:
  - Recent Activity (empty list + empty state)
  - Announcements/Notes (placeholder)
- Optional:
  - Chart placeholders (containers with empty-state text)

### Empty State
- Activity list shows empty state when no items.

---

## 2) Courses — Create/Manage Courses

### Purpose
Define programs (e.g., BSIT) used across the system.

### Top Bar
- Search (Course code/name)
- Status filter (Active/Inactive)
- Buttons: Clear, Refresh (UI only)

### Layout
**Left: DataGrid**
- Columns:
  - Course Code
  - Course Name
  - Status

**Right: Form**
- Fields:
  - Course Code
  - Course Name
  - Status
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Deactivate/Archive

### Empty State
- DataGrid overlay empty state.

---

## 3) Subjects — Create/Manage Subjects (Belongs to Course)

### Purpose
Subject masterlist tied to courses.

### Top Bar
- Course filter (ComboBox)
- Search (Code/Description)
- Status filter
- Buttons: Clear, Refresh

### Layout
**Left: DataGrid**
- Columns:
  - Subject Code
  - Description
  - Units
  - Course
  - Status

**Right: Form**
- Fields:
  - Subject Code
  - Description
  - Units
  - Course (ComboBox)
  - Status
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Archive/Deactivate

### Empty State
- DataGrid overlay empty state.

---

## 4) Sections — Create/Manage Sections

### Purpose
Define sections like BSIT-1A, BSIT-1B.

### Top Bar
- Course filter
- Year Level filter
- Search (Section name/code)
- Buttons: Clear, Refresh

### Layout
**Left: DataGrid**
- Columns:
  - Section Name/Code
  - Course
  - Year Level
  - Status

**Right: Form**
- Fields:
  - Section Name/Code
  - Course (ComboBox)
  - Year Level (ComboBox)
  - Status
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Deactivate

### Empty State
- DataGrid overlay empty state.

---

## 5) Rooms — Create/Manage Rooms

### Purpose
Define rooms for scheduling and filtering.

### Top Bar
- Search (Room name/code)
- Type filter (optional)
- Status filter
- Buttons: Clear, Refresh

### Layout
**Left: DataGrid**
- Columns:
  - Room Name/Code
  - Type (Lecture/Lab) (optional)
  - Capacity (optional)
  - Status

**Right: Form**
- Fields:
  - Room Name/Code
  - Type (optional)
  - Capacity (optional)
  - Status
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Deactivate

### Empty State
- DataGrid overlay empty state.

---

## 6) Students — Manage Students (Enroll per Subject Later)

### Purpose
Maintain student master records, filter by enrolled course (future). Enrollment module not implemented yet.

### Top Bar
- Search (Student ID / Name)
- Course filter (ComboBox)
- Status filter (Active/Graduated/Blocked)
- Buttons: Clear, Refresh

### Layout
**Left: DataGrid**
- Columns:
  - Student ID
  - Full Name
  - Course
  - Year Level
  - Status

**Right: Student Profile**
- Photo placeholder + Upload button (optional UI-only)
- Fields:
  - Student ID
  - First Name
  - Middle Name
  - Last Name
  - Sex
  - Birthdate
  - Address
  - Contact No
  - Guardian Name
  - Guardian Contact
  - Course (ComboBox)
  - Year Level (ComboBox)
  - Status (ComboBox)
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Mark as Graduated
  - Block/Unblock

**Optional Placeholder Panel**
- “Enrollment Summary” placeholder text

### Empty State
- DataGrid overlay empty state.

---

## 7) Teachers/Professors — Manage Teachers + Filter by Course/Day/Section/Room

### Purpose
Maintain teacher records and show assigned classes/schedule (future).

### Top Bar
- Search (Teacher ID / Name)
- Course filter
- Day filter
- Section filter
- Room filter
- Buttons: Clear, Refresh

### Layout
**Left: DataGrid**
- Columns:
  - Teacher ID
  - Full Name
  - Course/Department
  - Status

**Right: Teacher Profile**
- Photo placeholder + Upload button (optional UI-only)
- Fields:
  - Teacher ID
  - Full Name
  - Email (optional)
  - Contact No
  - Course/Department (ComboBox)
  - Status
- Buttons:
  - New
  - Edit
  - Save
  - Cancel
  - Block/Unblock

**Schedule Panel**
- Title: “Assigned Classes / Schedule”
- Columns:
  - Subject
  - Section
  - Room
  - Day
  - Time

### Empty State
- DataGrid overlay empty state for both teacher list and schedule list.

---

## 8) Scheduling — Tabs (Assignment Engine)

### Core Model
A **Class Schedule Entry**:
- Subject + Teacher + Section + Room + Day + StartTime + EndTime

Students will enroll per subject later; scheduling is about defining offerings/time slots.

### Tabs
#### Tab A: Class Scheduling (Main)
Top Filters:
- Course filter
- Section filter
- Room filter
- Day filter (optional)
- Buttons: Clear, Refresh

Schedule Grid Columns:
- Subject
- Teacher
- Section
- Room
- Day
- Start Time
- End Time

Form Fields:
- Subject (ComboBox, filtered by course)
- Teacher (ComboBox, filtered by course)
- Section (ComboBox)
- Room (ComboBox)
- Day (ComboBox)
- Start Time, End Time
Buttons:
- Add
- Update
- Remove
- Clear

Empty State:
- Grid overlay empty state.

#### Tab B: Teacher View
- Teacher selector (ComboBox/Search)
- Schedule list (same columns)
- Empty state overlay

#### Tab C: Section View
- Course + Section selectors
- Schedule list (same columns)
- Empty state overlay

---

## Recommended Build Order
1) Courses
2) Subjects
3) Sections
4) Rooms
5) Students
6) Teachers/Professors
7) Scheduling (Tabs)
8) Dashboard

---

## Completion Checklist (Per Page)
- [ ] Uses existing theme styles/brushes only (no new colors)
- [ ] No mock data added anywhere
- [ ] Empty-state overlay implemented for each grid/list
- [ ] ViewModel exists with empty collections + commands
- [ ] Page is navigable from AdminDashboardWindow host
- [ ] Compiles without missing references
