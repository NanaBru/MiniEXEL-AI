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

    public WorkbookView()
    {
        InitializeComponent();
        Pane1.CellEdited += async (s, sheetName) => await OnCellEdited(Pane2, sheetName);
        Pane2.CellEdited += async (s, sheetName) => await OnCellEdited(Pane1, sheetName);
    }

    private async Task OnCellEdited(SheetPaneView otherPane, string sheetName)
    {
        SetDirty(true);
        // Keep the other pane in sync if it happens to show the same sheet.
        await otherPane.RefreshIfShowingAsync(sheetName);
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
        foreach (var edit in plan.Edits)
        {
            var targetSheet = string.IsNullOrWhiteSpace(edit.Sheet) ? Pane1.CurrentSheetName : edit.Sheet;
            if (WorkbookService.TryApplyEditByAddress(_workbook, targetSheet, edit.Cell, edit.Value))
                applied++;
        }

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
