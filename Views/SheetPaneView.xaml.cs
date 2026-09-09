using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using MiniEXEL.Models;
using MiniEXEL.Services;

namespace MiniEXEL.Views;

// One self-contained "sheet viewer": formula bar, grid, image overlay and
// sheet tabs. WorkbookView hosts one or two of these side by side so the
// user can look at (and edit) two sheets of the same workbook at once.
public partial class SheetPaneView : UserControl
{
    private const double ColumnHeaderHeight = 26;
    private const double RowHeaderWidthPx = 45;
    private const double ColWidth = 90;
    private const double RowHeight = 22;

    // Raised whenever a cell edit is committed, so the host can mark the
    // workbook dirty and refresh any other pane showing the same sheet.
    public event EventHandler<string>? CellEdited;

    // Raised alongside CellEdited with enough detail (old + new value) for the
    // host to push an Undo/Redo entry. A single paste or edit can be a batch.
    public event EventHandler<List<EditUndoAction>>? EditsApplied;

    private int _sortColumnIndex = -1;
    private bool _sortAscending = true;
    private int _findMatchRow = -1, _findMatchCol = -1;

    private XLWorkbook? _workbook;
    private bool _supportsRichView;
    private string? _legacyPath;
    public string CurrentSheetName { get; private set; } = string.Empty;
    public object? GridWorkbookItemsSource => GridWorkbook.ItemsSource;

    private List<PictureSnapshot> _pictures = new();
    private double _hOffset;
    private double _vOffset;
    private readonly Dictionary<DataGridColumn, int> _columnIndexMap = new();
    private CellSnapshot? _selectedCell;

    public SheetPaneView()
    {
        InitializeComponent();
    }

    public async Task InitializeAsync(XLWorkbook? workbook, List<string> sheetNames, bool supportsRichView,
        string initialSheet, string? legacyPath, string? legacyNote)
    {
        _workbook = workbook;
        _supportsRichView = supportsRichView;
        _legacyPath = legacyPath;
        GridWorkbook.IsReadOnly = !supportsRichView;

        if (!supportsRichView)
        {
            TxtFormulaBar.Text = legacyNote ?? string.Empty;
            if (legacyPath != null) await LoadLegacyPlainViewAsync(legacyPath);
            return;
        }

        BuildSheetTabs(sheetNames, initialSheet);
        await LoadSheetAsync(initialSheet);
    }

    public async Task RefreshIfShowingAsync(string sheetName)
    {
        if (_supportsRichView && CurrentSheetName == sheetName)
            await LoadSheetAsync(sheetName);
    }

    public async Task SwitchToSheetAsync(string sheetName)
    {
        foreach (ToggleButton btn in SheetTabs.Items)
            btn.IsChecked = (string)btn.Content == sheetName;
        await LoadSheetAsync(sheetName);
    }

    public void RebuildTabs(List<string> sheetNames) => BuildSheetTabs(sheetNames, CurrentSheetName);

    private void BuildSheetTabs(List<string> sheetNames, string selected)
    {
        SheetTabs.ItemsSource = null;
        var buttons = new List<ToggleButton>();
        foreach (var name in sheetNames)
        {
            var btn = new ToggleButton
            {
                Content = name,
                Padding = new Thickness(12, 4, 12, 4),
                Margin = new Thickness(2),
                IsChecked = name == selected
            };
            btn.Click += async (s, e) =>
            {
                foreach (var b in buttons) b.IsChecked = ReferenceEquals(b, btn);
                await LoadSheetAsync(name);
            };
            buttons.Add(btn);
        }
        SheetTabs.ItemsSource = buttons;
    }

    private async Task LoadSheetAsync(string sheetName)
    {
        if (_workbook == null) return;
        CurrentSheetName = sheetName;
        TxtFormulaBar.Text = "Cargando hoja...";

        var snapshot = await Task.Run(() => WorkbookService.BuildSnapshot(_workbook, sheetName));
        RenderSnapshot(snapshot);
    }

    private async Task LoadLegacyPlainViewAsync(string path)
    {
        var preview = await Task.Run(() => ExcelPreviewService.LoadPreview(path));
        SheetTabs.ItemsSource = null;
        ImageOverlay.Children.Clear();
        _pictures.Clear();

        GridWorkbook.Columns.Clear();
        foreach (System.Data.DataColumn col in preview.Table.Columns)
        {
            GridWorkbook.Columns.Add(new DataGridTextColumn
            {
                Header = col.ColumnName,
                Binding = new System.Windows.Data.Binding($"[{col.ColumnName}]")
            });
        }
        GridWorkbook.ItemsSource = preview.Table.DefaultView;
        TxtFormulaBar.Text = $"{preview.Table.Rows.Count} filas mostradas" + (preview.Truncated ? " (archivo truncado por tamaño)" : "");
    }

    private void RenderSnapshot(SheetSnapshot snapshot)
    {
        GridWorkbook.Columns.Clear();
        _columnIndexMap.Clear();

        for (int c = 0; c < snapshot.ColCount; c++)
        {
            int colIndex = c;
            var column = new DataGridTemplateColumn
            {
                Width = ColWidth,
                ClipboardContentBinding = new System.Windows.Data.Binding($"Cells[{colIndex}].DisplayValue")
            };
            _columnIndexMap[column] = colIndex;

            // Clickable header: sorts the on-screen rows by this column (view
            // only — it never reorders rows in the actual .xlsx file).
            var headerBtn = new FrameworkElementFactory(typeof(Button));
            headerBtn.SetValue(Button.ContentProperty, ColumnLetter(c));
            headerBtn.SetValue(Button.BackgroundProperty, Brushes.Transparent);
            headerBtn.SetValue(Button.BorderThicknessProperty, new Thickness(0));
            headerBtn.SetValue(Button.PaddingProperty, new Thickness(0));
            headerBtn.SetValue(Button.CursorProperty, System.Windows.Input.Cursors.Hand);
            headerBtn.SetValue(Button.ToolTipProperty, "Ordenar por esta columna (solo en pantalla, no cambia el archivo)");
            headerBtn.AddHandler(Button.ClickEvent, new RoutedEventHandler((s, e) => SortByColumn(colIndex)));
            column.HeaderTemplate = new DataTemplate { VisualTree = headerBtn };

            // Each cell draws its own right/bottom border so grid lines stay
            // straight and continuous regardless of the cell's fill color —
            // relying on DataGrid's own grid lines let opaque cell backgrounds
            // paint over them inconsistently from one cell to the next.
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding($"Cells[{colIndex}].Background"));
            border.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8)));
            border.SetValue(Border.BorderThicknessProperty, new Thickness(0, 0, 1, 1));

            var text = new FrameworkElementFactory(typeof(TextBlock));
            text.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding($"Cells[{colIndex}].DisplayValue"));
            text.SetBinding(TextBlock.ForegroundProperty, new System.Windows.Data.Binding($"Cells[{colIndex}].Foreground"));
            text.SetBinding(TextBlock.FontWeightProperty, new System.Windows.Data.Binding($"Cells[{colIndex}].FontWeightValue"));
            text.SetValue(TextBlock.PaddingProperty, new Thickness(4, 2, 4, 2));
            text.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
            border.AppendChild(text);
            column.CellTemplate = new DataTemplate { VisualTree = border };

            var editBox = new FrameworkElementFactory(typeof(TextBox));
            // OneWay: we read the edited text straight off the TextBox in
            // CellEditEnding, so the binding never needs to write back. A
            // TwoWay/default binding to this read-only property throws
            // InvalidOperationException the moment WPF tries to attach a
            // write path to it — e.g. when pasting (Ctrl+V) into a cell.
            editBox.SetBinding(TextBox.TextProperty, new System.Windows.Data.Binding($"Cells[{colIndex}].RawEditText") { Mode = System.Windows.Data.BindingMode.OneWay });
            editBox.SetValue(TextBox.BorderThicknessProperty, new Thickness(0));
            editBox.SetValue(TextBox.PaddingProperty, new Thickness(4, 2, 4, 2));
            column.CellEditingTemplate = new DataTemplate { VisualTree = editBox };

            GridWorkbook.Columns.Add(column);
        }

        GridWorkbook.ItemsSource = snapshot.Rows;

        _pictures = snapshot.Pictures;
        RenderImages();

        var truncNote = snapshot.Truncated ? $" — vista limitada a {WorkbookService.MaxRows}x{WorkbookService.MaxCols} celdas" : "";
        TxtFormulaBar.Text = $"Hoja \"{snapshot.Name}\": {snapshot.RowCount} filas x {snapshot.ColCount} columnas{truncNote}";
    }

    private void RenderImages()
    {
        ImageOverlay.Children.Clear();
        foreach (var pic in _pictures)
        {
            var img = new System.Windows.Controls.Image
            {
                Source = pic.Image,
                Width = pic.WidthPx,
                Height = pic.HeightPx,
                Stretch = Stretch.Fill
            };
            Canvas.SetLeft(img, RowHeaderWidthPx + pic.AnchorColIndex * ColWidth - _hOffset);
            Canvas.SetTop(img, ColumnHeaderHeight + pic.AnchorRowIndex * RowHeight - _vOffset);
            ImageOverlay.Children.Add(img);
        }
    }

    private void BtnScrollTabsLeft_Click(object sender, RoutedEventArgs e) =>
        TabsScroll.ScrollToHorizontalOffset(TabsScroll.HorizontalOffset - 150);

    private void BtnScrollTabsRight_Click(object sender, RoutedEventArgs e) =>
        TabsScroll.ScrollToHorizontalOffset(TabsScroll.HorizontalOffset + 150);

    private void TabsScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        TabsScroll.ScrollToHorizontalOffset(TabsScroll.HorizontalOffset - e.Delta);
        e.Handled = true;
    }

    private void GridWorkbook_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        _hOffset = e.HorizontalOffset;
        _vOffset = e.VerticalOffset;
        if (_pictures.Count > 0) RenderImages();
    }

    private void GridWorkbook_CurrentCellChanged(object sender, EventArgs e)
    {
        var cellInfo = GridWorkbook.CurrentCell;
        if (cellInfo.Column == null || cellInfo.Item is not ExcelRowVm row) return;
        if (!_columnIndexMap.TryGetValue(cellInfo.Column, out int colIndex)) return;
        if (colIndex < 0 || colIndex >= row.Cells.Length) return;

        _selectedCell = row.Cells[colIndex];
        TxtNameBox.Text = _selectedCell.Address;
        TxtFormulaBar.Text = _selectedCell.RawEditText;
    }

    // DataGridTemplateColumn (used here so cells can carry per-cell colors)
    // doesn't support WPF's automatic clipboard paste like DataGridBoundColumn
    // does — and letting the default Paste command reach our read-only-bound
    // editor is what caused the crash pasting used to trigger. So Ctrl+V is
    // handled entirely by hand: split the clipboard like Excel does (rows by
    // newline, columns by tab) and write each piece starting at the active cell.
    private async void GridWorkbook_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.V || Keyboard.Modifiers != ModifierKeys.Control) return;
        e.Handled = true;
        if (!_supportsRichView || _workbook == null) return;

        string clip;
        try { clip = Clipboard.GetText(); }
        catch { return; }
        if (string.IsNullOrEmpty(clip)) return;

        var cellInfo = GridWorkbook.CurrentCell;
        if (cellInfo.Column == null || cellInfo.Item is not ExcelRowVm anchorRow) return;
        if (!_columnIndexMap.TryGetValue(cellInfo.Column, out int anchorColIndex)) return;

        var lines = clip.Replace("\r\n", "\n").TrimEnd('\n').Split('\n');
        int maxRow = GridWorkbook.Items.Count;
        int maxCol = _columnIndexMap.Count;
        bool anyApplied = false;

        var batch = new List<EditUndoAction>();
        var ws = _workbook.Worksheet(CurrentSheetName);

        for (int r = 0; r < lines.Length; r++)
        {
            var cols = lines[r].Split('\t');
            int targetRowNumber = anchorRow.RowNumber + r;
            if (targetRowNumber > maxRow) break; // stays within the currently loaded/bounded sheet area

            for (int c = 0; c < cols.Length; c++)
            {
                int targetColNumber = anchorRow.Cells[anchorColIndex].ColNumber + c;
                if (targetColNumber > maxCol) break;

                string oldValue = GetRawCellText(ws, targetRowNumber, targetColNumber);
                WorkbookService.ApplyEdit(_workbook, CurrentSheetName, targetRowNumber, targetColNumber, cols[c]);
                batch.Add(new EditUndoAction { SheetName = CurrentSheetName, RowNumber = targetRowNumber, ColNumber = targetColNumber, OldValue = oldValue, NewValue = cols[c] });
                anyApplied = true;
            }
        }

        if (!anyApplied) return;

        var sheetName = CurrentSheetName;
        await LoadSheetAsync(sheetName);
        CellEdited?.Invoke(this, sheetName);
        if (batch.Count > 0) EditsApplied?.Invoke(this, batch);
    }

    private static string GetRawCellText(IXLWorksheet ws, int row, int col)
    {
        var cell = ws.Cell(row, col);
        if (cell.HasFormula) return "=" + cell.FormulaA1;
        try { return cell.Value.ToString() ?? string.Empty; }
        catch { return cell.GetString(); }
    }

    private void GridWorkbook_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (!_supportsRichView || _workbook == null) return;
        if (e.EditAction != DataGridEditAction.Commit) return;
        if (e.Row.Item is not ExcelRowVm row) return;
        if (!_columnIndexMap.TryGetValue(e.Column, out int colIndex)) return;
        if (e.EditingElement is not TextBox tb) return;

        ApplyCellEdit(row.Cells[colIndex], tb.Text);
    }

    private void TxtFormulaBar_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || _selectedCell == null || !_supportsRichView) return;
        ApplyCellEdit(_selectedCell, TxtFormulaBar.Text);
        e.Handled = true;
    }

    private async void ApplyCellEdit(CellSnapshot cell, string newRawText)
    {
        if (_workbook == null) return;
        if (newRawText == cell.RawEditText) return;

        string oldValue = cell.RawEditText;
        WorkbookService.ApplyEdit(_workbook, CurrentSheetName, cell.RowNumber, cell.ColNumber, newRawText);

        var sheetName = CurrentSheetName;
        await LoadSheetAsync(sheetName);

        TxtNameBox.Text = cell.Address;
        TxtFormulaBar.Text = newRawText;

        CellEdited?.Invoke(this, sheetName);
        EditsApplied?.Invoke(this, new List<EditUndoAction>
        {
            new EditUndoAction { SheetName = sheetName, RowNumber = cell.RowNumber, ColNumber = cell.ColNumber, OldValue = oldValue, NewValue = newRawText }
        });
    }

    // Applies a batch of raw values directly (used by Undo/Redo). Does not
    // raise EditsApplied itself — the caller (WorkbookView) owns the stacks.
    public async Task ApplyRawBatchAsync(List<EditUndoAction> actions, bool useOldValue)
    {
        if (_workbook == null) return;
        string? sheetName = null;
        foreach (var action in actions)
        {
            WorkbookService.ApplyEdit(_workbook, action.SheetName, action.RowNumber, action.ColNumber, useOldValue ? action.OldValue : action.NewValue);
            sheetName = action.SheetName;
        }
        if (sheetName != null) await RefreshIfShowingAsync(sheetName);
    }

    // ---------------- View-only sort ----------------

    private void SortByColumn(int colIndex)
    {
        if (GridWorkbook.ItemsSource is not IEnumerable<ExcelRowVm> rows) return;

        _sortAscending = _sortColumnIndex == colIndex ? !_sortAscending : true;
        _sortColumnIndex = colIndex;

        var list = rows.ToList();
        list.Sort((a, b) =>
        {
            var va = colIndex < a.Cells.Length ? a.Cells[colIndex].DisplayValue : string.Empty;
            var vb = colIndex < b.Cells.Length ? b.Cells[colIndex].DisplayValue : string.Empty;
            int cmp;
            if (double.TryParse(va, out var da) && double.TryParse(vb, out var db))
                cmp = da.CompareTo(db);
            else
                cmp = string.Compare(va, vb, StringComparison.CurrentCultureIgnoreCase);
            return _sortAscending ? cmp : -cmp;
        });

        GridWorkbook.ItemsSource = list;
        TxtFormulaBar.Text = $"Ordenado por columna {ColumnLetter(colIndex)} ({(_sortAscending ? "ascendente" : "descendente")}) — solo en pantalla, no modifica el archivo.";
    }

    // ---------------- Find ----------------

    // Cycles to the next cell (row-major, wrapping) whose display text
    // contains the query. Returns false if nothing in the loaded sheet matches.
    public bool FindNext(string query)
    {
        if (GridWorkbook.ItemsSource is not IEnumerable<ExcelRowVm> rowsEnum || string.IsNullOrEmpty(query)) return false;
        var rows = rowsEnum.ToList();
        if (rows.Count == 0 || rows[0].Cells.Length == 0) return false;

        int colCount = rows[0].Cells.Length;
        int total = rows.Count * colCount;

        // Flat index of the last match (or -1 if this is the first search), so
        // "next" always resumes right after wherever we left off, wrapping around.
        int startIndex = _findMatchRow < 0 ? -1 : _findMatchRow * colCount + _findMatchCol;

        for (int step = 1; step <= total; step++)
        {
            int idx = (startIndex + step) % total;
            int r = idx / colCount;
            int c = idx % colCount;

            var cell = rows[r].Cells[c];
            if (cell.DisplayValue.Contains(query, StringComparison.CurrentCultureIgnoreCase))
            {
                _findMatchRow = r; _findMatchCol = c;
                SelectCell(rows[r], c);
                return true;
            }
        }
        return false;
    }

    public void ResetFind() { _findMatchRow = -1; _findMatchCol = -1; }

    private void SelectCell(ExcelRowVm row, int colIndex)
    {
        var column = GridWorkbook.Columns.FirstOrDefault(col => _columnIndexMap.TryGetValue(col, out var idx) && idx == colIndex);
        if (column == null) return;
        GridWorkbook.SelectedItem = row;
        GridWorkbook.CurrentCell = new DataGridCellInfo(row, column);
        GridWorkbook.ScrollIntoView(row, column);
    }

    // Replaces every match in the currently loaded sheet with newText (skips
    // formula cells so a running calculation is never silently overwritten).
    public async Task<int> ReplaceAllAsync(string query, string newText)
    {
        if (_workbook == null || GridWorkbook.ItemsSource is not IEnumerable<ExcelRowVm> rowsEnum || string.IsNullOrEmpty(query))
            return 0;

        var batch = new List<EditUndoAction>();
        foreach (var row in rowsEnum.ToList())
        {
            foreach (var cell in row.Cells)
            {
                if (cell.Formula != null) continue;
                if (!cell.DisplayValue.Contains(query, StringComparison.CurrentCultureIgnoreCase)) continue;

                string oldValue = cell.RawEditText;
                string replaced = System.Text.RegularExpressions.Regex.Replace(cell.DisplayValue, System.Text.RegularExpressions.Regex.Escape(query), newText, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                WorkbookService.ApplyEdit(_workbook, CurrentSheetName, cell.RowNumber, cell.ColNumber, replaced);
                batch.Add(new EditUndoAction { SheetName = CurrentSheetName, RowNumber = cell.RowNumber, ColNumber = cell.ColNumber, OldValue = oldValue, NewValue = replaced });
            }
        }

        if (batch.Count > 0)
        {
            var sheetName = CurrentSheetName;
            await LoadSheetAsync(sheetName);
            CellEdited?.Invoke(this, sheetName);
            EditsApplied?.Invoke(this, batch);
        }
        return batch.Count;
    }

    private static string ColumnLetter(int zeroBasedIndex)
    {
        int n = zeroBasedIndex + 1;
        string result = string.Empty;
        while (n > 0)
        {
            int rem = (n - 1) % 26;
            result = (char)('A' + rem) + result;
            n = (n - 1) / 26;
        }
        return result;
    }

    public string DumpForAi()
    {
        if (GridWorkbook.ItemsSource is not IEnumerable<ExcelRowVm> rows) return string.Empty;
        var sb = new System.Text.StringBuilder();
        foreach (var row in rows)
        {
            foreach (var cell in row.Cells)
            {
                if (string.IsNullOrWhiteSpace(cell.DisplayValue) && cell.Formula == null) continue;
                sb.Append(cell.Address).Append('=').Append(cell.Formula != null ? "=" + cell.Formula : cell.DisplayValue).Append("; ");
            }
        }
        return sb.ToString();
    }

    public void Clear()
    {
        _workbook = null;
        _pictures.Clear();
        ImageOverlay.Children.Clear();
        GridWorkbook.ItemsSource = null;
        GridWorkbook.Columns.Clear();
        _columnIndexMap.Clear();
        SheetTabs.ItemsSource = null;
        TxtFormulaBar.Text = string.Empty;
        TxtNameBox.Text = string.Empty;
        CurrentSheetName = string.Empty;
        _sortColumnIndex = -1;
        ResetFind();
    }
}
