Attribute VB_Name = "TaskDashboard"
Option Explicit

' ============================================================================
' TaskDashboard — BuildDashboard macro
' Reads the "Task Export" sheet and rebuilds a formatted "Dashboard" sheet:
'   * KPI tiles (total / completed / in progress / overdue / completion %)
'   * Tasks by owner table
'   * Tasks by status table
'   * Overdue task list (red) and due-in-next-7-days list (amber)
'
' How to use: open Task-Dashboard.xlsm, enable macros, press Alt+F8,
' select BuildDashboard, click Run. Re-run any time — the Dashboard sheet
' is deleted and rebuilt from the current Task Export data.
'
' 100% local. No network calls, no external files. by Muhammad Bilal.
' ============================================================================

Public Sub BuildDashboard()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim src As Worksheet
    Dim dash As Worksheet
    Dim todayDate As Date
    Dim lastRow As Long
    Dim n As Long
    Dim i As Long, j As Long, k As Long

    Dim ids() As String
    Dim titles() As String
    Dim projects() As String
    Dim owners() As String
    Dim statuses() As String
    Dim priorities() As String
    Dim dues() As Date

    Dim totalCount As Long
    Dim completedCount As Long
    Dim inProgressCount As Long
    Dim blockedCount As Long
    Dim notStartedCount As Long
    Dim overdueCount As Long
    Dim dueSoonCount As Long

    Dim ownerNames() As String
    Dim ownerTotal() As Long
    Dim ownerDone() As Long
    Dim ownerOverdue() As Long
    Dim ownerCount As Long

    Dim statusNames(1 To 4) As String
    Dim statusCount(1 To 4) As Long

    Dim overIdx() As Long
    Dim soonIdx() As Long
    Dim tmp As Long
    Dim r As Long

    Set wb = ThisWorkbook
    Set src = wb.Worksheets("Task Export")
    todayDate = Date

    ' ---------- read tasks (scan down column A until empty) ----------
    lastRow = 1
    Do While Trim(CStr(src.Cells(lastRow + 1, 1).Value)) <> ""
        lastRow = lastRow + 1
    Loop
    n = lastRow - 1
    If n <= 0 Then
        MsgBox "No tasks found in 'Task Export'.", vbExclamation, "BuildDashboard"
        Exit Sub
    End If

    ReDim ids(1 To n)
    ReDim titles(1 To n)
    ReDim projects(1 To n)
    ReDim owners(1 To n)
    ReDim statuses(1 To n)
    ReDim priorities(1 To n)
    ReDim dues(1 To n)

    For i = 1 To n
        ids(i) = CStr(src.Cells(i + 1, 1).Value)
        titles(i) = CStr(src.Cells(i + 1, 2).Value)
        projects(i) = CStr(src.Cells(i + 1, 3).Value)
        owners(i) = CStr(src.Cells(i + 1, 4).Value)
        statuses(i) = CStr(src.Cells(i + 1, 5).Value)
        priorities(i) = CStr(src.Cells(i + 1, 6).Value)
        dues(i) = CDate(src.Cells(i + 1, 7).Value)
    Next i

    ' ---------- KPI counts ----------
    statusNames(1) = "Not Started"
    statusNames(2) = "In Progress"
    statusNames(3) = "Completed"
    statusNames(4) = "Blocked"

    For i = 1 To n
        totalCount = totalCount + 1
        Select Case statuses(i)
            Case "Completed": completedCount = completedCount + 1
            Case "In Progress": inProgressCount = inProgressCount + 1
            Case "Blocked": blockedCount = blockedCount + 1
            Case "Not Started": notStartedCount = notStartedCount + 1
        End Select
        For j = 1 To 4
            If statuses(i) = statusNames(j) Then statusCount(j) = statusCount(j) + 1
        Next j
        If statuses(i) <> "Completed" And dues(i) < todayDate Then overdueCount = overdueCount + 1
        If statuses(i) <> "Completed" And dues(i) >= todayDate And dues(i) <= todayDate + 7 Then dueSoonCount = dueSoonCount + 1
    Next i

    ' ---------- per-owner aggregation ----------
    ownerCount = 0
    ReDim ownerNames(1 To n)
    ReDim ownerTotal(1 To n)
    ReDim ownerDone(1 To n)
    ReDim ownerOverdue(1 To n)
    For i = 1 To n
        k = 0
        For j = 1 To ownerCount
            If ownerNames(j) = owners(i) Then k = j
        Next j
        If k = 0 Then
            ownerCount = ownerCount + 1
            k = ownerCount
            ownerNames(k) = owners(i)
        End If
        ownerTotal(k) = ownerTotal(k) + 1
        If statuses(i) = "Completed" Then ownerDone(k) = ownerDone(k) + 1
        If statuses(i) <> "Completed" And dues(i) < todayDate Then ownerOverdue(k) = ownerOverdue(k) + 1
    Next i

    ' ---------- collect + sort overdue / due-soon indices by due date ----------
    ReDim overIdx(1 To n)
    ReDim soonIdx(1 To n)
    overdueCount = 0
    dueSoonCount = 0
    For i = 1 To n
        If statuses(i) <> "Completed" And dues(i) < todayDate Then
            overdueCount = overdueCount + 1
            overIdx(overdueCount) = i
        End If
        If statuses(i) <> "Completed" And dues(i) >= todayDate And dues(i) <= todayDate + 7 Then
            dueSoonCount = dueSoonCount + 1
            soonIdx(dueSoonCount) = i
        End If
    Next i
    ' Bubble sort (small lists; deterministic, no Range.Sort dependency)
    For i = 1 To overdueCount - 1
        For j = i + 1 To overdueCount
            If dues(overIdx(j)) < dues(overIdx(i)) Then
                tmp = overIdx(i): overIdx(i) = overIdx(j): overIdx(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To dueSoonCount - 1
        For j = i + 1 To dueSoonCount
            If dues(soonIdx(j)) < dues(soonIdx(i)) Then
                tmp = soonIdx(i): soonIdx(i) = soonIdx(j): soonIdx(j) = tmp
            End If
        Next j
    Next i

    ' ---------- rebuild Dashboard sheet ----------
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Dashboard").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    Set dash = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    dash.Name = "Dashboard"

    Call SetColWidth(dash, "A", 14)
    Call SetColWidth(dash, "B", 42)
    Call SetColWidth(dash, "C", 26)
    Call SetColWidth(dash, "D", 16)
    Call SetColWidth(dash, "E", 13)
    Call SetColWidth(dash, "F", 13)
    Call SetColWidth(dash, "G", 13)
    Call SetColWidth(dash, "H", 13)
    Call SetColWidth(dash, "I", 13)
    Call SetColWidth(dash, "J", 13)

    ' Title
    dash.Range("A1:J1").Merge
    dash.Range("A1").Value = "TASK DASHBOARD"
    dash.Range("A1").Font.Name = "Calibri"
    dash.Range("A1").Font.Size = 20
    dash.Range("A1").Font.Bold = True
    dash.Range("A1").Font.Color = RGB(255, 255, 255)
    dash.Range("A1").Interior.Color = RGB(14, 76, 129)
    dash.Range("A1").HorizontalAlignment = xlCenter
    dash.Range("A1").VerticalAlignment = xlCenter
    Call SetRowHeight(dash, 1, 36)

    dash.Range("A2:J2").Merge
    dash.Range("A2").Value = "Generated " & Format(todayDate, "yyyy-mm-dd") & " from 'Task Export'  |  Re-run BuildDashboard any time  |  by Muhammad Bilal"
    dash.Range("A2").Font.Name = "Calibri"
    dash.Range("A2").Font.Size = 10
    dash.Range("A2").Font.Italic = True
    dash.Range("A2").Font.Color = RGB(90, 90, 90)
    dash.Range("A2").HorizontalAlignment = xlCenter

    ' ---------- KPI tiles (row 4 label, row 5 value) ----------
    Dim tileLabels(1 To 5) As String
    Dim tileValues(1 To 5) As Variant
    Dim tileColors(1 To 5) As Long
    tileLabels(1) = "TOTAL TASKS"
    tileLabels(2) = "COMPLETED"
    tileLabels(3) = "IN PROGRESS"
    tileLabels(4) = "OVERDUE"
    tileLabels(5) = "COMPLETION %"
    tileValues(1) = totalCount
    tileValues(2) = completedCount
    tileValues(3) = inProgressCount
    tileValues(4) = overdueCount
    If totalCount > 0 Then tileValues(5) = completedCount / totalCount Else tileValues(5) = 0
    tileColors(1) = RGB(14, 76, 129)
    tileColors(2) = RGB(33, 128, 74)
    tileColors(3) = RGB(190, 120, 10)
    tileColors(4) = RGB(180, 30, 30)
    tileColors(5) = RGB(0, 120, 140)

    For k = 1 To 5
        Dim c1 As Long, c2 As Long
        c1 = (k - 1) * 2 + 1
        c2 = c1 + 1
        Dim labelCell As Range, valueCell As Range
        Set labelCell = dash.Range(dash.Cells(4, c1), dash.Cells(4, c2))
        Set valueCell = dash.Range(dash.Cells(5, c1), dash.Cells(5, c2))
        labelCell.Merge
        valueCell.Merge
        labelCell.Value = tileLabels(k)
        labelCell.Font.Name = "Calibri"
        labelCell.Font.Size = 10
        labelCell.Font.Bold = True
        labelCell.Font.Color = RGB(255, 255, 255)
        labelCell.Interior.Color = tileColors(k)
        labelCell.HorizontalAlignment = xlCenter
        valueCell.Value = tileValues(k)
        valueCell.Font.Name = "Calibri"
        valueCell.Font.Size = 22
        valueCell.Font.Bold = True
        valueCell.Font.Color = tileColors(k)
        valueCell.HorizontalAlignment = xlCenter
        If k = 5 Then valueCell.NumberFormat = "0%"
        Call StyleBox(dash.Range(dash.Cells(4, c1), dash.Cells(5, c2)))
    Next k
    Call SetRowHeight(dash, 4, 22)
    Call SetRowHeight(dash, 5, 34)

    ' ---------- Tasks by Owner ----------
    r = 7
    dash.Range("A" & r & ":E" & r).Merge
    dash.Range("A" & r).Value = "TASKS BY OWNER"
    Call SectionHeader(dash.Range("A" & r & ":E" & r))
    r = r + 1
    dash.Cells(r, 1).Value = "Owner"
    dash.Cells(r, 2).Value = "Total"
    dash.Cells(r, 3).Value = "Completed"
    dash.Cells(r, 4).Value = "Overdue"
    dash.Cells(r, 5).Value = "Completion %"
    Call TableHeader(dash.Range("A" & r & ":E" & r))
    r = r + 1
    For k = 1 To ownerCount
        dash.Cells(r, 1).Value = ownerNames(k)
        dash.Cells(r, 2).Value = ownerTotal(k)
        dash.Cells(r, 3).Value = ownerDone(k)
        dash.Cells(r, 4).Value = ownerOverdue(k)
        If ownerTotal(k) > 0 Then dash.Cells(r, 5).Value = ownerDone(k) / ownerTotal(k) Else dash.Cells(r, 5).Value = 0
        dash.Cells(r, 5).NumberFormat = "0%"
        Call TableRow(dash.Range("A" & r & ":E" & r), (r Mod 2 = 0))
        r = r + 1
    Next k

    ' ---------- Tasks by Status ----------
    r = r + 1
    dash.Range("A" & r & ":C" & r).Merge
    dash.Range("A" & r).Value = "TASKS BY STATUS"
    Call SectionHeader(dash.Range("A" & r & ":C" & r))
    r = r + 1
    dash.Cells(r, 1).Value = "Status"
    dash.Cells(r, 2).Value = "Count"
    dash.Cells(r, 3).Value = "% of Total"
    Call TableHeader(dash.Range("A" & r & ":C" & r))
    r = r + 1
    For k = 1 To 4
        dash.Cells(r, 1).Value = statusNames(k)
        dash.Cells(r, 2).Value = statusCount(k)
        If totalCount > 0 Then dash.Cells(r, 3).Value = statusCount(k) / totalCount Else dash.Cells(r, 3).Value = 0
        dash.Cells(r, 3).NumberFormat = "0%"
        Call TableRow(dash.Range("A" & r & ":C" & r), (r Mod 2 = 0))
        r = r + 1
    Next k

    ' ---------- Overdue list ----------
    r = r + 1
    dash.Range("A" & r & ":F" & r).Merge
    dash.Range("A" & r).Value = "OVERDUE TASKS  —  action required  (" & overdueCount & ")"
    Call SectionHeader(dash.Range("A" & r & ":F" & r))
    r = r + 1
    dash.Cells(r, 1).Value = "Task ID"
    dash.Cells(r, 2).Value = "Title"
    dash.Cells(r, 3).Value = "Project"
    dash.Cells(r, 4).Value = "Owner"
    dash.Cells(r, 5).Value = "Due Date"
    dash.Cells(r, 6).Value = "Days Overdue"
    Call TableHeader(dash.Range("A" & r & ":F" & r))
    r = r + 1
    For k = 1 To overdueCount
        i = overIdx(k)
        dash.Cells(r, 1).Value = ids(i)
        dash.Cells(r, 2).Value = titles(i)
        dash.Cells(r, 3).Value = projects(i)
        dash.Cells(r, 4).Value = owners(i)
        dash.Cells(r, 5).Value = dues(i)
        dash.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
        dash.Cells(r, 6).Value = todayDate - dues(i)
        Call AlertRow(dash.Range("A" & r & ":F" & r), RGB(255, 205, 205))
        r = r + 1
    Next k
    If overdueCount = 0 Then
        dash.Range("A" & r & ":F" & r).Merge
        dash.Range("A" & r).Value = "No overdue tasks — nice work."
        r = r + 1
    End If

    ' ---------- Due soon list ----------
    r = r + 1
    dash.Range("A" & r & ":F" & r).Merge
    dash.Range("A" & r).Value = "DUE IN NEXT 7 DAYS  (" & dueSoonCount & ")"
    Call SectionHeader(dash.Range("A" & r & ":F" & r))
    r = r + 1
    dash.Cells(r, 1).Value = "Task ID"
    dash.Cells(r, 2).Value = "Title"
    dash.Cells(r, 3).Value = "Project"
    dash.Cells(r, 4).Value = "Owner"
    dash.Cells(r, 5).Value = "Due Date"
    dash.Cells(r, 6).Value = "Days Left"
    Call TableHeader(dash.Range("A" & r & ":F" & r))
    r = r + 1
    For k = 1 To dueSoonCount
        i = soonIdx(k)
        dash.Cells(r, 1).Value = ids(i)
        dash.Cells(r, 2).Value = titles(i)
        dash.Cells(r, 3).Value = projects(i)
        dash.Cells(r, 4).Value = owners(i)
        dash.Cells(r, 5).Value = dues(i)
        dash.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
        dash.Cells(r, 6).Value = dues(i) - todayDate
        Call AlertRow(dash.Range("A" & r & ":F" & r), RGB(255, 235, 180))
        r = r + 1
    Next k
    If dueSoonCount = 0 Then
        dash.Range("A" & r & ":F" & r).Merge
        dash.Range("A" & r).Value = "Nothing due in the next 7 days."
        r = r + 1
    End If

    ' Footer
    r = r + 1
    dash.Range("A" & r & ":F" & r).Merge
    dash.Range("A" & r).Value = "Built by the BuildDashboard macro (TaskDashboard.bas)  |  100% local — no network calls  |  by Muhammad Bilal"
    dash.Range("A" & r).Font.Size = 9
    dash.Range("A" & r).Font.Italic = True
    dash.Range("A" & r).Font.Color = RGB(100, 100, 100)

    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "BuildDashboard failed: " & Err.Description, vbCritical, "BuildDashboard"
End Sub

' ---------- formatting helpers (errors swallowed: formatting must never kill the build) ----------

Private Sub SetColWidth(ws As Worksheet, colLetter As String, w As Double)
    On Error Resume Next
    ws.Columns(colLetter).ColumnWidth = w
End Sub

Private Sub SetRowHeight(ws As Worksheet, rowNum As Long, h As Double)
    On Error Resume Next
    ws.Rows(rowNum).RowHeight = h
End Sub

Private Sub StyleBox(rng As Range)
    On Error Resume Next
    Dim edge As Long
    For edge = 7 To 12
        rng.Borders(edge).LineStyle = xlContinuous
        rng.Borders(edge).Weight = xlThin
        rng.Borders(edge).Color = RGB(160, 160, 160)
    Next edge
End Sub

Private Sub SectionHeader(rng As Range)
    On Error Resume Next
    rng.Font.Name = "Calibri"
    rng.Font.Size = 12
    rng.Font.Bold = True
    rng.Font.Color = RGB(255, 255, 255)
    rng.Interior.Color = RGB(14, 76, 129)
    rng.VerticalAlignment = xlCenter
    rng.RowHeight = 24
End Sub

Private Sub TableHeader(rng As Range)
    On Error Resume Next
    Dim c As Range
    For Each c In rng.Cells
        c.Font.Name = "Calibri"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = RGB(255, 255, 255)
        c.Interior.Color = RGB(90, 110, 130)
        c.HorizontalAlignment = xlCenter
        c.VerticalAlignment = xlCenter
    Next c
    rng.RowHeight = 20
    Call StyleBox(rng)
End Sub

Private Sub TableRow(rng As Range, shade As Boolean)
    On Error Resume Next
    Dim c As Range
    For Each c In rng.Cells
        c.Font.Name = "Calibri"
        c.Font.Size = 10
        c.VerticalAlignment = xlCenter
        If shade Then c.Interior.Color = RGB(240, 244, 248)
    Next c
    Call StyleBox(rng)
End Sub

Private Sub AlertRow(rng As Range, fillColor As Long)
    On Error Resume Next
    Dim c As Range
    For Each c In rng.Cells
        c.Font.Name = "Calibri"
        c.Font.Size = 10
        c.VerticalAlignment = xlCenter
        c.Interior.Color = fillColor
    Next c
    Call StyleBox(rng)
End Sub
