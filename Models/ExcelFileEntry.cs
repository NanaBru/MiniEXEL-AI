using System;

namespace MiniEXEL.Models
{
    public class ExcelFileEntry
    {
        public string FullPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public long SizeBytes { get; set; }
        public DateTime LastModifiedUtc { get; set; }
        public DateTime LastSeenUtc { get; set; }

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
    }
}
