using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using MiniEXEL.Models;

namespace MiniEXEL.Services
{
    // Persists the detected-file list as JSON in %LOCALAPPDATA%\MiniEXEL.
    // Dedup is enforced by normalized full path so re-scans never create duplicates,
    // and entries whose file no longer exists on disk are pruned on load to keep
    // the cache small (no unbounded growth over time).
    public class CacheService
    {
        private readonly string _cacheDir;
        private readonly string _cacheFile;
        private readonly string _foldersFile;

        public CacheService()
        {
            _cacheDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MiniEXEL");
            Directory.CreateDirectory(_cacheDir);
            _cacheFile = Path.Combine(_cacheDir, "cache.json");
            _foldersFile = Path.Combine(_cacheDir, "folders.json");
        }

        public List<ExcelFileEntry> Load()
        {
            if (!File.Exists(_cacheFile)) return new List<ExcelFileEntry>();

            try
            {
                var json = File.ReadAllText(_cacheFile);
                var list = JsonSerializer.Deserialize<List<ExcelFileEntry>>(json) ?? new List<ExcelFileEntry>();

                // Dedup defensively in case an older cache file had duplicates.
                var deduped = list
                    .GroupBy(e => e.Key)
                    .Select(g => g.OrderByDescending(e => e.LastSeenUtc).First())
                    .Where(e => File.Exists(e.FullPath)) // prune stale entries
                    .OrderByDescending(e => e.LastSeenUtc)
                    .ToList();

                return deduped;
            }
            catch
            {
                // Corrupt cache file: start fresh instead of crashing the app.
                return new List<ExcelFileEntry>();
            }
        }

        public void Save(IEnumerable<ExcelFileEntry> entries)
        {
            var deduped = entries
                .GroupBy(e => e.Key)
                .Select(g => g.OrderByDescending(e => e.LastSeenUtc).First())
                .OrderByDescending(e => e.LastSeenUtc)
                .ToList();

            var json = JsonSerializer.Serialize(deduped, new JsonSerializerOptions { WriteIndented = false });
            File.WriteAllText(_cacheFile, json);
        }

        public List<string> LoadExtraFolders()
        {
            if (!File.Exists(_foldersFile)) return new List<string>();
            try
            {
                var json = File.ReadAllText(_foldersFile);
                return JsonSerializer.Deserialize<List<string>>(json) ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        public void SaveExtraFolders(IEnumerable<string> folders)
        {
            var distinct = folders.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var json = JsonSerializer.Serialize(distinct);
            File.WriteAllText(_foldersFile, json);
        }
    }
}
