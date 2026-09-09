using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MiniEXEL.Models;

namespace MiniEXEL.Services
{
    // Scans the Windows "Recent" shell folder (files opened before) plus any
    // user-added folders, and returns Excel files found. Uses streaming
    // EnumerateFiles instead of GetFiles to avoid buffering large arrays in RAM.
    public static class FileScanner
    {
        private static readonly string[] ExcelExtensions = { ".xlsx", ".xlsm", ".xls" };

        public static IEnumerable<ExcelFileEntry> ScanRecent()
        {
            string recentDir = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            if (!Directory.Exists(recentDir)) yield break;

            IEnumerable<string> lnkFiles;
            try
            {
                lnkFiles = Directory.EnumerateFiles(recentDir, "*.lnk", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                yield break;
            }

            foreach (var lnk in lnkFiles)
            {
                string? target = ShortcutResolver.ResolveTarget(lnk);
                if (string.IsNullOrWhiteSpace(target)) continue;
                if (!ExcelExtensions.Contains(Path.GetExtension(target), StringComparer.OrdinalIgnoreCase)) continue;

                var entry = TryBuildEntry(target);
                if (entry != null) yield return entry;
            }
        }

        public static IEnumerable<ExcelFileEntry> ScanFolder(string folder, int maxDepth = 3)
        {
            if (!Directory.Exists(folder)) yield break;

            var stack = new Stack<(string dir, int depth)>();
            stack.Push((folder, 0));

            while (stack.Count > 0)
            {
                var (dir, depth) = stack.Pop();

                IEnumerable<string> files;
                try { files = Directory.EnumerateFiles(dir); }
                catch { continue; }

                foreach (var f in files)
                {
                    if (!ExcelExtensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)) continue;
                    // Skip Excel lock files like ~$file.xlsx
                    if (Path.GetFileName(f).StartsWith("~$")) continue;

                    var entry = TryBuildEntry(f);
                    if (entry != null) yield return entry;
                }

                if (depth >= maxDepth) continue;

                IEnumerable<string> subdirs;
                try { subdirs = Directory.EnumerateDirectories(dir); }
                catch { continue; }

                foreach (var sub in subdirs)
                    stack.Push((sub, depth + 1));
            }
        }

        private static ExcelFileEntry? TryBuildEntry(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                if (!fi.Exists) return null;

                return new ExcelFileEntry
                {
                    FullPath = fi.FullName,
                    FileName = fi.Name,
                    SizeBytes = fi.Length,
                    LastModifiedUtc = fi.LastWriteTimeUtc,
                    LastSeenUtc = DateTime.UtcNow
                };
            }
            catch
            {
                return null;
            }
        }
    }
}
