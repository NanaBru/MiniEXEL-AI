using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using MiniEXEL.Models;
using MiniEXEL.Services;

namespace MiniEXEL.Views;

public partial class WorkbookView : UserControl
{
    public event EventHandler? BackRequested;

    private XLWorkbook? _workbook;
    private string _currentPath = string.Empty;
    private List<string> _sheetNames = new();
    private bool _isDirty;
    private bool _supportsRichView;
    private bool _splitView;

    private readonly AiSettingsService _aiSettingsService = new();
    private readonly List<ChatTurn> _chatHistory = new();

    private readonly Stack<List<EditUndoAction>> _undoStack = new();
    private readonly Stack<List<EditUndoAction>> _redoStack = new();
    private bool _applyingUndo;

    public WorkbookView()
    {
        InitializeComponent();
        Pane1.CellEdited += async (s, sheetName) => await OnCellEdited(Pane2, sheetName);
        Pane2.CellEdited += async (s, sheetName) => await OnCellEdited(Pane1, sheetName);
        Pane1.EditsApplied += (s, batch) => OnEditsApplied(batch);
        Pane2.EditsApplied += (s, batch) => OnEditsApplied(batch);
    }

    private async Task OnCellEdited(SheetPaneView otherPane, string sheetName)
    {
        SetDirty(true);
        // Keep the other pane in sync if it happens to show the same sheet.
        await otherPane.RefreshIfShowingAsync(sheetName);
    }

    private void OnEditsApplied(List<EditUndoAction> batch)
    {
        if (_applyingUndo || batch.Count == 0) return;
        _undoStack.Push(batch);
        _redoStack.Clear();
        UpdateUndoRedoButtons();
    }

    private void UpdateUndoRedoButtons()
    {
        BtnUndo.IsEnabled = _undoStack.Count > 0;
        BtnRedo.IsEnabled = _redoStack.Count > 0;
    }

    private async void BtnUndo_Click(object sender, RoutedEventArgs e)
    {
        if (_undoStack.Count == 0) return;
        var batch = _undoStack.Pop();
        _applyingUndo = true;
        try
        {
            await Pane1.ApplyRawBatchAsync(batch, useOldValue: true);
            await Pane2.ApplyRawBatchAsync(batch, useOldValue: true);
        }
        finally { _applyingUndo = false; }
        _redoStack.Push(batch);
        SetDirty(true);
        UpdateUndoRedoButtons();
    }

    private async void BtnRedo_Click(object sender, RoutedEventArgs e)
    {
        if (_redoStack.Count == 0) return;
        var batch = _redoStack.Pop();
        _applyingUndo = true;
        try
        {
            await Pane1.ApplyRawBatchAsync(batch, useOldValue: false);
            await Pane2.ApplyRawBatchAsync(batch, useOldValue: false);
        }
        finally { _applyingUndo = false; }
        _undoStack.Push(batch);
        SetDirty(true);
        UpdateUndoRedoButtons();
    }

    // ---------------- Keyboard shortcuts ----------------

    private void Root_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        bool ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        if (!ctrl)
        {
            if (e.Key == Key.Escape && FindBar.Visibility == Visibility.Visible)
            {
                BtnCloseFind_Click(sender, e);
                e.Handled = true;
            }
            return;
        }

        switch (e.Key)
        {
            case Key.Z: BtnUndo_Click(sender, e); e.Handled = true; break;
            case Key.Y: BtnRedo_Click(sender, e); e.Handled = true; break;
            case Key.S:
                if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) BtnSaveAs_Click(sender, e);
                else BtnSave_Click(sender, e);
                e.Handled = true;
                break;
            case Key.F: BtnFind_Click(sender, e); e.Handled = true; break;
        }
    }

    public async Task OpenFileAsync(string path)
    {
        CloseWorkbook();

        _currentPath = path;
        TxtFileName.Text = Path.GetFileName(path);
        TxtNote.Text = string.Empty;
        TxtSaveStatus.Text = string.Empty;
        SetDirty(false);

        var result = await Task.Run(() => WorkbookService.Open(path));
        _supportsRichView = result.SupportsRichView;
        _sheetNames = result.SheetNames;

        if (!result.SupportsRichView)
        {
            TxtNote.Text = result.LimitationNote;
            await Pane1.InitializeAsync(null, _sheetNames, false, _sheetNames.First(), path, result.LimitationNote);
            return;
        }

        _workbook = result.Workbook;
        await Pane1.InitializeAsync(_workbook, _sheetNames, true, _sheetNames.First(), null, null);
    }

    private async void BtnSplit_Click(object sender, RoutedEventArgs e)
    {
        if (_workbook == null || _sheetNames.Count == 0) return;

        _splitView = !_splitView;

        if (_splitView)
        {
            SplitterCol.Width = new GridLength(6);
            Pane2Col.Width = new GridLength(1, GridUnitType.Star);
            PaneSplitter.Width = 6;
            PaneSplitter.Visibility = Visibility.Visible;
            Pane2.Visibility = Visibility.Visible;

            // Default the second pane to the next sheet if there is one, so the
            // user immediately sees two different sheets side by side.
            var secondSheet = _sheetNames.Count > 1 ? _sheetNames[1] : _sheetNames[0];
            await Pane2.InitializeAsync(_workbook, _sheetNames, true, secondSheet, null, null);

            BtnSplit.Content = "◧ Una hoja";
        }
        else
        {
            SplitterCol.Width = new GridLength(0);
            Pane2Col.Width = new GridLength(0);
            PaneSplitter.Width = 0;
            PaneSplitter.Visibility = Visibility.Collapsed;
            Pane2.Visibility = Visibility.Collapsed;
            Pane2.Clear();

            BtnSplit.Content = "▤ Ver dos hojas";
        }
    }

    private void SetDirty(bool dirty)
    {
        _isDirty = dirty;
        TxtDirty.Text = dirty ? "● sin guardar" : string.Empty;
    }

    private async void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        if (_workbook == null || !_supportsRichView)
        {
            TxtSaveStatus.Text = "Este formato no se puede guardar desde aquí.";
            return;
        }

        TxtSaveStatus.Text = "Guardando...";
        try
        {
            await Task.Run(() => _workbook.Save());
            SetDirty(false);
            TxtSaveStatus.Text = "Guardado.";
        }
        catch (Exception ex)
        {
            TxtSaveStatus.Text = $"Error al guardar: {ex.Message}";
        }
    }

    private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
    {
        if (_workbook == null || !_supportsRichView)
        {
            TxtSaveStatus.Text = "Este formato no se puede guardar desde aquí.";
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title = "Guardar copia como",
            Filter = "Libro de Excel (*.xlsx)|*.xlsx",
            FileName = Path.GetFileNameWithoutExtension(_currentPath) + " - copia.xlsx",
            InitialDirectory = Path.GetDirectoryName(_currentPath)
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            _workbook.SaveAs(dlg.FileName);
            _currentPath = dlg.FileName;
            TxtFileName.Text = Path.GetFileName(_currentPath);
            SetDirty(false);
            TxtSaveStatus.Text = "Guardado como " + Path.GetFileName(_currentPath) + ".";
        }
        catch (Exception ex)
        {
            TxtSaveStatus.Text = $"Error al guardar: {ex.Message}";
        }
    }

    private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
    {
        if (Pane1.GridWorkbookItemsSource is not System.Collections.IEnumerable rowsEnum)
        {
            TxtSaveStatus.Text = "No hay una hoja cargada para exportar.";
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title = "Exportar hoja a CSV",
            Filter = "CSV (*.csv)|*.csv",
            FileName = Path.GetFileNameWithoutExtension(_currentPath) + " - " + Pane1.CurrentSheetName + ".csv"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var sb = new StringBuilder();
            foreach (ExcelRowVm row in rowsEnum)
            {
                sb.AppendLine(string.Join(",", row.Cells.Select(c => CsvEscape(c.DisplayValue))));
            }
            File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true)); // BOM so Excel opens accented text correctly
            TxtSaveStatus.Text = "Exportado a " + Path.GetFileName(dlg.FileName) + ".";
        }
        catch (Exception ex)
        {
            TxtSaveStatus.Text = $"Error al exportar: {ex.Message}";
        }
    }

    private static string CsvEscape(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        return value;
    }

    // ---------------- Find & replace ----------------

    // Cross-sheet find cursor: which sheet/row/col the last match was on, so
    // "Siguiente" resumes from there and wraps around the whole workbook.
    private string? _findSheet;
    private int _findRow = -1, _findCol = -1;

    private void BtnFind_Click(object sender, RoutedEventArgs e)
    {
        FindBar.Visibility = FindBar.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        if (FindBar.Visibility == Visibility.Visible)
        {
            ResetFindCursor();
            TxtFind.Focus();
        }
    }

    private void ResetFindCursor()
    {
        _findSheet = null; _findRow = -1; _findCol = -1;
        Pane1.ResetFind();
    }

    private void BtnCloseFind_Click(object sender, RoutedEventArgs e)
    {
        FindBar.Visibility = Visibility.Collapsed;
        TxtFindStatus.Text = string.Empty;
    }

    private void TxtFind_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { BtnFindNext_Click(sender, e); e.Handled = true; }
    }

    private async void BtnFindNext_Click(object sender, RoutedEventArgs e)
    {
        var query = TxtFind.Text;
        if (string.IsNullOrEmpty(query) || _workbook == null) return;

        bool found = await FindNextAcrossSheetsAsync(query);
        TxtFindStatus.Text = found
            ? $"Encontrado en \"{_findSheet}\"."
            : "Sin más resultados en todo el libro.";
        if (!found) ResetFindCursor();
    }

    // Searches every sheet of the workbook (starting right after the last
    // match, wrapping around) for the query, switches the primary pane to
    // whichever sheet it lands on, and selects the matching cell there.
    private async Task<bool> FindNextAcrossSheetsAsync(string query)
    {
        if (_workbook == null || _sheetNames.Count == 0) return false;

        int startSheetIdx = _findSheet != null ? _sheetNames.IndexOf(_findSheet) : 0;
        if (startSheetIdx < 0) startSheetIdx = 0;

        for (int s = 0; s < _sheetNames.Count; s++)
        {
            int sheetIdx = (startSheetIdx + s) % _sheetNames.Count;
            var sheetName = _sheetNames[sheetIdx];
            var ws = _workbook.Worksheet(sheetName);
            var used = ws.RangeUsed();
            if (used == null) continue;

            int maxR = Math.Min(used.RowCount(), WorkbookService.MaxRows);
            int maxC = Math.Min(used.ColumnCount(), WorkbookService.MaxCols);
            bool sameSheetAsLastMatch = sheetName == _findSheet;

            for (int r = 1; r <= maxR; r++)
            {
                for (int c = 1; c <= maxC; c++)
                {
                    // Skip everything up to (and including) where we left off.
                    if (sameSheetAsLastMatch && (r < _findRow || (r == _findRow && c <= _findCol))) continue;

                    var cell = ws.Cell(r, c);
                    string text;
                    try { text = cell.Value.ToString() ?? string.Empty; }
                    catch { text = cell.GetString(); }

                    if (text.Contains(query, StringComparison.CurrentCultureIgnoreCase))
                    {
                        _findSheet = sheetName; _findRow = r; _findCol = c;
                        if (Pane1.CurrentSheetName != sheetName)
                            await Pane1.SwitchToSheetAsync(sheetName);
                        Pane1.SelectCellByAddress(cell.Address.ToString());
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private async void BtnReplaceAll_Click(object sender, RoutedEventArgs e)
    {
        var query = TxtFind.Text;
        if (string.IsNullOrEmpty(query)) return;

        // Replace stays scoped to the sheet currently shown, on purpose — a
        // silent workbook-wide rewrite is too easy to regret.
        int count = await Pane1.ReplaceAllAsync(query, TxtReplace.Text);
        TxtFindStatus.Text = count > 0 ? $"{count} reemplazo(s) hecho(s) en \"{Pane1.CurrentSheetName}\"." : "Sin coincidencias en esta hoja.";
        if (count > 0) SetDirty(true);
    }

    public void CloseWorkbook()
    {
        _workbook?.Dispose();
        _workbook = null;
        _sheetNames.Clear();
        Pane1.Clear();
        Pane2.Clear();

        _splitView = false;
        SplitterCol.Width = new GridLength(0);
        Pane2Col.Width = new GridLength(0);
        PaneSplitter.Visibility = Visibility.Collapsed;
        Pane2.Visibility = Visibility.Collapsed;
        BtnSplit.Content = "▤ Ver dos hojas";

        FindBar.Visibility = Visibility.Collapsed;
        _findSheet = null; _findRow = -1; _findCol = -1;
        _undoStack.Clear();
        _redoStack.Clear();
        UpdateUndoRedoButtons();

        _chatHistory.Clear();
        ChatMessages.Items.Clear();
        AiPanel.Visibility = Visibility.Collapsed;
        SetDirty(false);
    }

    private void BtnBack_Click(object sender, RoutedEventArgs e)
    {
        if (_isDirty)
        {
            var result = MessageBox.Show("Hay cambios sin guardar. ¿Guardar antes de salir?",
                "Cambios sin guardar", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Cancel) return;
            if (result == MessageBoxResult.Yes)
            {
                try { _workbook?.Save(); } catch { /* fall through, still navigate back */ }
            }
        }

        CloseWorkbook();
        BackRequested?.Invoke(this, EventArgs.Empty);
    }

    // ---------------- AI assistant ----------------

    private void BtnAi_Click(object sender, RoutedEventArgs e)
    {
        AiPanel.Visibility = AiPanel.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
    }

    private void BtnAiSettings_Click(object sender, RoutedEventArgs e)
    {
        var current = _aiSettingsService.Load();
        var dlg = new AiSettingsWindow(current) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            _aiSettingsService.Save(dlg.Result);
            TxtAiStatus.Text = "Configuración guardada.";
        }
    }

    private void TxtChatInput_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            _ = SendChatAsync();
        }
    }

    private void BtnSendChat_Click(object sender, RoutedEventArgs e) => _ = SendChatAsync();

    private async Task SendChatAsync()
    {
        var question = TxtChatInput.Text.Trim();
        if (string.IsNullOrEmpty(question)) return;

        var settings = _aiSettingsService.Load();
        if (!settings.IsConfigured)
        {
            TxtAiStatus.Text = "Configura el proveedor de IA primero (⚙ Configurar).";
            return;
        }
        if (_workbook == null)
        {
            TxtAiStatus.Text = "Abre una hoja con soporte completo (.xlsx) para usar el asistente.";
            return;
        }

        AddChatBubble(question, isUser: true);
        TxtChatInput.Text = string.Empty;
        TxtAiStatus.Text = "Pensando...";
        BtnSendChat.IsEnabled = false;

        try
        {
            if (_chatHistory.Count == 0)
                _chatHistory.Add(new ChatTurn { Role = "system", Content = BuildSystemPrompt() });
            else
                _chatHistory[0] = new ChatTurn { Role = "system", Content = BuildSystemPrompt() }; // keep sheet dump fresh

            _chatHistory.Add(new ChatTurn { Role = "user", Content = question });

            var reply = await AiChatService.SendAsync(settings, _chatHistory);
            _chatHistory.Add(new ChatTurn { Role = "assistant", Content = reply });

            AddChatBubble(reply, isUser: false);

            var plan = AiChatService.TryExtractEditPlan(reply);
            if (plan != null && (plan.NewSheets.Count > 0 || plan.Edits.Count > 0))
                AddApplyPlanButton(plan);

            TxtAiStatus.Text = string.Empty;
        }
        catch (Exception ex)
        {
            TxtAiStatus.Text = $"Error: {ex.Message}";
        }
        finally
        {
            BtnSendChat.IsEnabled = true;
        }
    }

    private string BuildSystemPrompt()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Eres un asistente que ayuda a analizar y editar un libro Excel dentro de la app MiniEXEL.");
        sb.AppendLine("Responde en español, breve y concreto.");
        sb.AppendLine("Puedes proponer cambios a la hoja (no solo responder consultas): editar celdas existentes, o crear una hoja nueva y llenarla con datos.");
        sb.AppendLine("Cuando propongas cambios, agrega al final de tu respuesta un bloque JSON con este formato exacto:");
        sb.AppendLine("```json");
        sb.AppendLine("{\"new_sheets\": [\"NombreHojaNueva\"], \"edits\": [{\"sheet\": \"NombreHojaNueva\", \"cell\": \"A1\", \"value\": \"Total\"}, {\"cell\": \"D2\", \"value\": \"=B2*C2\"}]}");
        sb.AppendLine("```");
        sb.AppendLine("\"new_sheets\" es opcional (lista de hojas a crear si no existen). \"edits\" es la lista de celdas a escribir.");
        sb.AppendLine("En cada edit, \"sheet\" es opcional: si lo omites, se aplica a la hoja que se está mostrando actualmente. Si editas una hoja nueva, indica su nombre en \"sheet\".");
        sb.AppendLine("\"value\" empieza con \"=\" para fórmulas, o es el valor literal para texto/números.");
        sb.AppendLine();
        sb.AppendLine("Hojas existentes en el libro: " + string.Join(", ", _sheetNames));
        sb.AppendLine($"Hoja mostrada actualmente en el panel principal: \"{Pane1.CurrentSheetName}\"");
        sb.AppendLine("Contenido de esa hoja (dirección=valor):");
        sb.AppendLine(Pane1.DumpForAi());

        if (Pane2.Visibility == Visibility.Visible && !string.IsNullOrEmpty(Pane2.CurrentSheetName))
        {
            sb.AppendLine($"Hoja mostrada en el segundo panel: \"{Pane2.CurrentSheetName}\"");
            sb.AppendLine("Contenido de esa hoja (dirección=valor):");
            sb.AppendLine(Pane2.DumpForAi());
        }

        return sb.ToString();
    }

    private void AddChatBubble(string text, bool isUser)
    {
        var border = new Border
        {
            Background = isUser ? new SolidColorBrush(Color.FromRgb(0x1B, 0x4B, 0x33)) : new SolidColorBrush(Color.FromRgb(0x1E, 0x26, 0x22)),
            BorderBrush = isUser ? (Brush)Application.Current.Resources["AccentDim"] : (Brush)Application.Current.Resources["BorderSoft"],
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(10),
            Margin = new Thickness(isUser ? 40 : 0, 4, isUser ? 0 : 40, 4),
            HorizontalAlignment = isUser ? HorizontalAlignment.Right : HorizontalAlignment.Left
        };
        border.Child = new TextBlock
        {
            Text = text,
            TextWrapping = TextWrapping.Wrap,
            Foreground = (Brush)Application.Current.Resources["TextPrimary"]
        };

        ((System.Collections.IList)ChatMessages.Items).Add(border);
        ChatScroll.ScrollToEnd();
    }

    private void AddApplyPlanButton(AiEditPlan plan)
    {
        int total = plan.Edits.Count + plan.NewSheets.Count;
        var btn = new Button
        {
            Content = $"Aplicar {total} cambio(s) sugerido(s) por la IA",
            Style = (Style)Application.Current.Resources["Win11ButtonAccent"],
            Margin = new Thickness(0, 4, 0, 8),
            HorizontalAlignment = HorizontalAlignment.Left
        };
        btn.Click += (s, e) => ApplyAiPlan(plan, btn);

        ((System.Collections.IList)ChatMessages.Items).Add(btn);
        ChatScroll.ScrollToEnd();
    }

    private async void ApplyAiPlan(AiEditPlan plan, Button triggerButton)
    {
        if (_workbook == null) return;

        bool sheetsChanged = false;
        foreach (var sheetName in plan.NewSheets)
        {
            if (WorkbookService.EnsureSheetExists(_workbook, sheetName))
                sheetsChanged = true;
        }

        if (sheetsChanged)
        {
            _sheetNames = _workbook.Worksheets.Select(w => w.Name).ToList();
            Pane1.RebuildTabs(_sheetNames);
            if (Pane2.Visibility == Visibility.Visible) Pane2.RebuildTabs(_sheetNames);
        }

        int applied = 0;
        var undoBatch = new List<EditUndoAction>();
        foreach (var edit in plan.Edits)
        {
            var targetSheet = string.IsNullOrWhiteSpace(edit.Sheet) ? Pane1.CurrentSheetName : edit.Sheet;
            var ws = _workbook.Worksheet(targetSheet);
            string oldValue;
            try
            {
                var cell = ws.Cell(edit.Cell);
                oldValue = cell.HasFormula ? "=" + cell.FormulaA1 : (cell.Value.ToString() ?? string.Empty);
            }
            catch { oldValue = string.Empty; }

            if (WorkbookService.TryApplyEditByAddress(_workbook, targetSheet, edit.Cell, edit.Value))
            {
                applied++;
                var addr = ws.Cell(edit.Cell).Address;
                undoBatch.Add(new EditUndoAction { SheetName = targetSheet, RowNumber = addr.RowNumber, ColNumber = addr.ColumnNumber, OldValue = oldValue, NewValue = edit.Value });
            }
        }
        if (undoBatch.Count > 0) OnEditsApplied(undoBatch);

        SetDirty(true);
        triggerButton.IsEnabled = false;
        triggerButton.Content = $"Aplicado(s) {applied}/{plan.Edits.Count} cambio(s)" + (plan.NewSheets.Count > 0 ? $" + {plan.NewSheets.Count} hoja(s) nueva(s)" : "");

        await Pane1.RefreshIfShowingAsync(Pane1.CurrentSheetName);
        if (Pane2.Visibility == Visibility.Visible)
            await Pane2.RefreshIfShowingAsync(Pane2.CurrentSheetName);

        // If the AI created a sheet with edits targeting it, jump the primary
        // pane to that new sheet so the user immediately sees the result.
        var firstNewSheetWithData = plan.NewSheets.FirstOrDefault(s => plan.Edits.Any(e => e.Sheet == s));
        if (firstNewSheetWithData != null)
            await Pane1.SwitchToSheetAsync(firstNewSheetWithData);
    }
}
