using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Interop;
using Microsoft.Win32;
using MiniEXEL.Models;
using MiniEXEL.Services;

namespace MiniEXEL;

public partial class MainWindow : Window
{
    private readonly CacheService _cache = new();
    private readonly ObservableCollection<ExcelFileEntry> _files = new();
    private readonly ICollectionView _filesView;
    private List<string> _extraFolders = new();
    private string _searchText = string.Empty;

    public MainWindow()
    {
        InitializeComponent();

        GridFiles.ItemsSource = _files;
        _filesView = CollectionViewSource.GetDefaultView(_files);
        _filesView.Filter = FilterFiles;

        WorkbookHost.BackRequested += (s, e) => ShowListView();

        SourceInitialized += MainWindow_SourceInitialized;

        LoadFromCacheAndRescan();
        StartRamMonitor();
    }

    // ---------------- Maximize-covers-taskbar fix ----------------
    // A borderless WindowChrome window maximized via WindowState alone expands
    // to the full monitor bounds (ignoring the taskbar), which is why content
    // like the AI panel ended up rendered underneath it. Hooking WM_GETMINMAXINFO
    // and clamping to the monitor's work area is the correct, general fix (unlike
    // padding the root grid, it also accounts for taskbar position/size and
    // multi-monitor setups).
    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        var handle = new WindowInteropHelper(this).Handle;
        HwndSource.FromHwnd(handle)?.AddHook(WindowProc);
    }

    private const int WM_GETMINMAXINFO = 0x0024;

    private IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_GETMINMAXINFO)
        {
            WmGetMinMaxInfo(hwnd, lParam);
            handled = true;
        }
        return IntPtr.Zero;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public int dwFlags;
    }

    [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr handle, int flags);
    [DllImport("user32.dll")] private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

    private const int MONITOR_DEFAULTTONEAREST = 2;

    private static void WmGetMinMaxInfo(IntPtr hwnd, IntPtr lParam)
    {
        var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);

        var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
        if (monitor != IntPtr.Zero)
        {
            var info = new MONITORINFO { cbSize = Marshal.SizeOf<MONITORINFO>() };
            GetMonitorInfo(monitor, ref info);

            var workArea = info.rcWork;
            var monitorArea = info.rcMonitor;

            mmi.ptMaxPosition.X = Math.Abs(workArea.Left - monitorArea.Left);
            mmi.ptMaxPosition.Y = Math.Abs(workArea.Top - monitorArea.Top);
            mmi.ptMaxSize.X = Math.Abs(workArea.Right - workArea.Left);
            mmi.ptMaxSize.Y = Math.Abs(workArea.Bottom - workArea.Top);
            mmi.ptMaxTrackSize.X = mmi.ptMaxSize.X;
            mmi.ptMaxTrackSize.Y = mmi.ptMaxSize.Y;
        }

        Marshal.StructureToPtr(mmi, lParam, true);
    }

    private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            BtnMaximizeRestore_Click(sender, e);
            return;
        }
        if (e.ButtonState == System.Windows.Input.MouseButtonState.Pressed)
            DragMove();
    }

    private void BtnMinimize_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

    private void BtnMaximizeRestore_Click(object sender, RoutedEventArgs e)
    {
        WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

    private async void OpenWorkbookView(ExcelFileEntry entry)
    {
        ListHost.Visibility = Visibility.Collapsed;
        WorkbookHost.Visibility = Visibility.Visible;
        await WorkbookHost.OpenFileAsync(entry.FullPath);
    }

    private void ShowListView()
    {
        WorkbookHost.Visibility = Visibility.Collapsed;
        ListHost.Visibility = Visibility.Visible;
    }

    private void GridFiles_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (GridFiles.SelectedItem is ExcelFileEntry entry)
            OpenWorkbookView(entry);
    }

    private void BtnOpenWorkbook_Click(object sender, RoutedEventArgs e)
    {
        if (GridFiles.SelectedItem is ExcelFileEntry entry)
            OpenWorkbookView(entry);
    }

    private bool FilterFiles(object obj)
    {
        if (string.IsNullOrWhiteSpace(_searchText)) return true;
        if (obj is not ExcelFileEntry entry) return false;
        return entry.FileName.Contains(_searchText, StringComparison.OrdinalIgnoreCase);
    }

    private void LoadFromCacheAndRescan()
    {
        _extraFolders = _cache.LoadExtraFolders();

        // Show cached results immediately (fast startup, no scanning wait).
        var cached = _cache.Load();
        ReplaceFiles(cached);
        TxtStatus.Text = $"{_files.Count} archivo(s) en caché. Actualizando...";

        RescanAsync();
    }

    private async void RescanAsync()
    {
        var extras = _extraFolders.ToList();

        var found = await System.Threading.Tasks.Task.Run(() =>
        {
            var results = new Dictionary<string, ExcelFileEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in FileScanner.ScanRecent())
                results[entry.Key] = entry;

            foreach (var folder in extras)
                foreach (var entry in FileScanner.ScanFolder(folder))
                    results[entry.Key] = entry;

            return results.Values.ToList();
        });

        // Merge with existing cache: keep LastSeenUtc fresh for files still found,
        // but don't drop entries the user manually kept if they are still on disk.
        var merged = new Dictionary<string, ExcelFileEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in _files) merged[e.Key] = e;
        foreach (var e in found) merged[e.Key] = e; // fresh scan wins (updates LastSeenUtc, size, date)

        var finalList = merged.Values
            .Where(e => System.IO.File.Exists(e.FullPath))
            .OrderByDescending(e => e.LastModifiedUtc)
            .ToList();

        ReplaceFiles(finalList);
        _cache.Save(finalList);
        TxtStatus.Text = $"{_files.Count} archivo(s) detectado(s).";
    }

    private void ReplaceFiles(IEnumerable<ExcelFileEntry> entries)
    {
        _files.Clear();
        foreach (var e in entries.OrderByDescending(x => x.LastModifiedUtc))
            _files.Add(e);
    }

    private void BtnRefresh_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Actualizando...";
        RescanAsync();
    }

    private void BtnAddFolder_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = "Selecciona una carpeta con archivos Excel" };
        if (dlg.ShowDialog() == true)
        {
            if (!_extraFolders.Contains(dlg.FolderName, StringComparer.OrdinalIgnoreCase))
            {
                _extraFolders.Add(dlg.FolderName);
                _cache.SaveExtraFolders(_extraFolders);
            }
            TxtStatus.Text = "Escaneando carpeta...";
            RescanAsync();
        }
    }

    private void BtnOpenExcel_Click(object sender, RoutedEventArgs e)
    {
        if (GridFiles.SelectedItem is not ExcelFileEntry entry) return;
        try
        {
            Process.Start(new ProcessStartInfo(entry.FullPath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"No se pudo abrir el archivo:\n{ex.Message}", "Error",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void BtnRemove_Click(object sender, RoutedEventArgs e)
    {
        if (GridFiles.SelectedItem is not ExcelFileEntry entry) return;
        _files.Remove(entry);
        _cache.Save(_files);
        TxtStatus.Text = $"{_files.Count} archivo(s) en la lista.";
    }

    private void TxtSearch_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        _searchText = TxtSearch.Text;
        _filesView.Refresh();
    }

    private static readonly string[] ExcelExtensions = { ".xlsx", ".xlsm", ".xls" };

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        bool hasExcel = e.Data.GetDataPresent(DataFormats.FileDrop) &&
            ((string[])e.Data.GetData(DataFormats.FileDrop)!).Any(HasExcelExtensionOrIsFolder);

        e.Effects = hasExcel ? DragDropEffects.Copy : DragDropEffects.None;
        DropOverlay.Visibility = hasExcel ? Visibility.Visible : Visibility.Collapsed;
        e.Handled = true;
    }

    private void Window_DragLeave(object sender, DragEventArgs e)
    {
        DropOverlay.Visibility = Visibility.Collapsed;
    }

    private static bool HasExcelExtensionOrIsFolder(string path)
    {
        if (System.IO.Directory.Exists(path)) return true;
        return ExcelExtensions.Contains(System.IO.Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        DropOverlay.Visibility = Visibility.Collapsed;
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var dropped = (string[])e.Data.GetData(DataFormats.FileDrop)!;
        var added = new List<ExcelFileEntry>();

        foreach (var path in dropped)
        {
            if (System.IO.Directory.Exists(path))
            {
                added.AddRange(FileScanner.ScanFolder(path));

                // Remember the folder so future scans keep picking it up too.
                if (!_extraFolders.Contains(path, StringComparer.OrdinalIgnoreCase))
                    _extraFolders.Add(path);
            }
            else if (System.IO.File.Exists(path) &&
                     ExcelExtensions.Contains(System.IO.Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            {
                var fi = new System.IO.FileInfo(path);
                added.Add(new ExcelFileEntry
                {
                    FullPath = fi.FullName,
                    FileName = fi.Name,
                    SizeBytes = fi.Length,
                    LastModifiedUtc = fi.LastWriteTimeUtc,
                    LastSeenUtc = DateTime.UtcNow
                });
            }
        }

        if (added.Count == 0) return;

        _cache.SaveExtraFolders(_extraFolders);

        var merged = new Dictionary<string, ExcelFileEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _files) merged[entry.Key] = entry;
        foreach (var entry in added) merged[entry.Key] = entry; // dropped files refresh existing entries too

        var finalList = merged.Values.OrderByDescending(x => x.LastModifiedUtc).ToList();
        ReplaceFiles(finalList);
        _cache.Save(finalList);

        var first = added.First();
        var match = _files.FirstOrDefault(f => f.Key == first.Key);
        if (match != null) GridFiles.SelectedItem = match;

        TxtStatus.Text = $"{added.Count} archivo(s) agregado(s) por arrastre. {_files.Count} en total.";
    }

    private void GridFiles_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (GridFiles.SelectedItem is ExcelFileEntry entry)
            TxtStatus.Text = $"Seleccionado: {entry.FileName}  ·  doble clic para abrirlo.";
    }

    // ---------------- Sidebar navigation ----------------

    private void NavInicio_Click(object sender, RoutedEventArgs e) => TxtSearch.Text = string.Empty;

    private void NavRecientes_Click(object sender, RoutedEventArgs e) => RescanAsync();

    private void NavFavoritos_Click(object sender, RoutedEventArgs e)
    {
        TxtStatus.Text = "Favoritos: próximamente.";
    }

    private void NavExplorar_Click(object sender, RoutedEventArgs e) => BtnAddFolder_Click(sender, e);

    private void NavConfig_Click(object sender, RoutedEventArgs e)
    {
        var settingsService = new AiSettingsService();
        var dlg = new Views.AiSettingsWindow(settingsService.Load()) { Owner = this };
        if (dlg.ShowDialog() == true)
            settingsService.Save(dlg.Result);
    }

    private void DropZone_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "Seleccionar archivo Excel",
            Filter = "Archivos Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls",
            Multiselect = true
        };
        if (dlg.ShowDialog() != true) return;

        var added = new List<ExcelFileEntry>();
        foreach (var path in dlg.FileNames)
        {
            var fi = new System.IO.FileInfo(path);
            if (!fi.Exists) continue;
            added.Add(new ExcelFileEntry
            {
                FullPath = fi.FullName,
                FileName = fi.Name,
                SizeBytes = fi.Length,
                LastModifiedUtc = fi.LastWriteTimeUtc,
                LastSeenUtc = DateTime.UtcNow
            });
        }
        if (added.Count == 0) return;

        var merged = new Dictionary<string, ExcelFileEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _files) merged[entry.Key] = entry;
        foreach (var entry in added) merged[entry.Key] = entry;

        var finalList = merged.Values.OrderByDescending(x => x.LastModifiedUtc).ToList();
        ReplaceFiles(finalList);
        _cache.Save(finalList);
        TxtStatus.Text = $"{added.Count} archivo(s) agregado(s). {_files.Count} en total.";
    }

    // ---------------- RAM usage indicator ----------------

    private readonly System.Windows.Threading.DispatcherTimer _ramTimer = new() { Interval = TimeSpan.FromSeconds(3) };

    private void StartRamMonitor()
    {
        UpdateRamUsage();
        _ramTimer.Tick += (s, e) => UpdateRamUsage();
        _ramTimer.Start();
    }

    private void UpdateRamUsage()
    {
        double mb = Process.GetCurrentProcess().WorkingSet64 / (1024.0 * 1024.0);
        TxtRamUsage.Text = $"{mb:0} MB";
        RamBar.Width = Math.Clamp(mb / 2, 6, 260); // rough visual scale, not a hard limit
    }
}
