using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace MiniEXEL.Models
{
    // A snapshot of one Excel cell's display value, formula and style. Implements
    // INotifyPropertyChanged so an in-place edit (grid, formula bar, or AI assistant)
    // refreshes the bound cell on screen without rebuilding the whole sheet.
    public class CellSnapshot : INotifyPropertyChanged
    {
        private string _displayValue = string.Empty;
        private string? _formula;

        public string Address { get; set; } = string.Empty;
        public int RowNumber { get; set; } // 1-based Excel row
        public int ColNumber { get; set; } // 1-based Excel column

        public string DisplayValue
        {
            get => _displayValue;
            set
            {
                if (_displayValue == value) return;
                _displayValue = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RawEditText));
            }
        }

        public string? Formula
        {
            get => _formula;
            set
            {
                if (_formula == value) return;
                _formula = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RawEditText));
            }
        }

        public SolidColorBrush Background { get; set; } = Brushes.Transparent;
        public SolidColorBrush Foreground { get; set; } = Brushes.Black;
        public FontWeight FontWeightValue { get; set; } = FontWeights.Normal;

        // Raw text as Excel's formula bar would show it (editable form).
        public string RawEditText => Formula != null ? "=" + Formula : DisplayValue;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
