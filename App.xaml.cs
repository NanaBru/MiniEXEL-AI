using System.Configuration;
using System.Data;
using System.Text;
using System.Windows;

namespace MiniEXEL;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        // ExcelDataReader needs this to read legacy .xls files (and some
        // non-UTF8 encoded workbooks) — without it it throws NotSupportedException.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        DispatcherUnhandledException += (s, ex) =>
        {
            MessageBox.Show(ex.Exception.ToString(), "Error inesperado",
                MessageBoxButton.OK, MessageBoxImage.Error);
            ex.Handled = true;
        };
    }
}

