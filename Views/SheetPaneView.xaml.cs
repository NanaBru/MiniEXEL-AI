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

    private XLWorkbook? _workbook;
    private bool _supportsRichView;
    private string? _legacyPath;
    public string CurrentSheetName { get; private set; } = string.Empty;

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
                Header = ColumnLetter(c),
                Width = ColWidth,
                ClipboardContentBinding = new System.Windows.Data.Binding($"Cells[{colIndex}].DisplayValue")
            };
            _columnIndexMap[column] = colIndex;

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

        for (int r = 0; r < lines.Length; r++)
        {
            var cols = lines[r].Split('\t');
            int targetRowNumber = anchorRow.RowNumber + r;
            if (targetRowNumber > maxRow) break; // stays within the currently loaded/bounded sheet area

            for (int c = 0; c < cols.Length; c++)
            {
                int targetColNumber = anchorRow.Cells[anchorColIndex].ColNumber + c;
                if (targetColNumber > maxCol) break;

                WorkbookService.ApplyEdit(_workbook, CurrentSheetName, targetRowNumber, targetColNumber, cols[c]);
                anyApplied = true;
            }
        }

        if (!anyApplied) return;

        var sheetName = CurrentSheetName;
        await LoadSheetAsync(sheetName);
        CellEdited?.Invoke(this, sheetName);
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

        WorkbookService.ApplyEdit(_workbook, CurrentSheetName, cell.RowNumber, cell.ColNumber, newRawText);

        var sheetName = CurrentSheetName;
        await LoadSheetAsync(sheetName);

        TxtNameBox.Text = cell.Address;
        TxtFormulaBar.Text = newRawText;

        CellEdited?.Invoke(this, sheetName);
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
    }
}
