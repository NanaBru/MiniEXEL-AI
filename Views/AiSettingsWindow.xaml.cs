using System.Windows;
using System.Windows.Controls;
using MiniEXEL.Models;

namespace MiniEXEL.Views;

public partial class AiSettingsWindow : Window
{
    public AiSettings Result { get; private set; }
    private bool _initializing = true;

    public AiSettingsWindow(AiSettings current)
    {
        InitializeComponent();
        Result = current;

        foreach (ComboBoxItem item in CmbProvider.Items)
        {
            if ((string)item.Tag == current.Provider) { CmbProvider.SelectedItem = item; break; }
        }
        if (CmbProvider.SelectedItem == null) CmbProvider.SelectedIndex = 0;

        TxtBaseUrl.Text = current.BaseUrl;
        TxtModel.Text = current.Model;
        TxtApiKey.Password = current.ApiKey;
        _initializing = false;
    }

    private void CmbProvider_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_initializing) return;
        if (CmbProvider.SelectedItem is not ComboBoxItem item) return;

        // Prefill a sensible base URL when switching to OpenRouter; leave it
        // blank for "Custom" so the user points it at their own provider.
        if ((string)item.Tag == "OpenRouter" && string.IsNullOrWhiteSpace(TxtBaseUrl.Text))
            TxtBaseUrl.Text = "https://openrouter.ai/api/v1/chat/completions";
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        var provider = (string)((ComboBoxItem)CmbProvider.SelectedItem).Tag;

        Result = new AiSettings
        {
            Provider = provider,
            BaseUrl = TxtBaseUrl.Text.Trim(),
            Model = TxtModel.Text.Trim(),
            ApiKey = TxtApiKey.Password
        };

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
