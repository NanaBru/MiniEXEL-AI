using System;
using System.IO;
using System.Text.Json;
using MiniEXEL.Models;

namespace MiniEXEL.Services
{
    // Stores AI provider settings (provider, base URL, model, API key) in
    // %LOCALAPPDATA%\MiniEXEL\ai_settings.json. Stored in plain text — same
    // trust boundary as any local desktop app's saved credentials; the API key
    // never leaves the machine except in requests to the configured provider.
    public class AiSettingsService
    {
        private readonly string _file;

        public AiSettingsService()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MiniEXEL");
            Directory.CreateDirectory(dir);
            _file = Path.Combine(dir, "ai_settings.json");
        }

        public AiSettings Load()
        {
            if (!File.Exists(_file)) return new AiSettings();
            try
            {
                var json = File.ReadAllText(_file);
                return JsonSerializer.Deserialize<AiSettings>(json) ?? new AiSettings();
            }
            catch
            {
                return new AiSettings();
            }
        }

        public void Save(AiSettings settings)
        {
            var json = JsonSerializer.Serialize(settings);
            File.WriteAllText(_file, json);
        }
    }
}
