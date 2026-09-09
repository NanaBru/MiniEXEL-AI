using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace MiniEXEL.Models
{
    public class ExcelFileEntry : INotifyPropertyChanged
    {
        private bool _isFavorite;

        public string FullPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public long SizeBytes { get; set; }
        public DateTime LastModifiedUtc { get; set; }
        public DateTime LastSeenUtc { get; set; }

        public bool IsFavorite
        {
            get => _isFavorite;
            set
            {
                if (_isFavorite == value) return;
                _isFavorite = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FavoriteGlyph));
            }
        }

        // Segoe MDL2 Assets glyphs: filled star (E735) vs outline star (E734)
        public string FavoriteGlyph => IsFavorite ? "" : "";

        public string SizeDisplay
        {
            get
            {
                double kb = SizeBytes / 1024.0;
                if (kb < 1024) return $"{kb:0.#} KB";
                return $"{kb / 1024.0:0.#} MB";
            }
        }

        // Unique key used for dedup: normalized path (case-insensitive on Windows)
        public string Key => FullPath.Trim().ToLowerInvariant();

        public string FolderPath => System.IO.Path.GetDirectoryName(FullPath) ?? string.Empty;

        public string ModifiedDisplay => LastModifiedUtc.ToLocalTime().ToString("dd/MM/yy HH:mm");

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
