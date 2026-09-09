using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniEXEL.Models;

namespace MiniEXEL.Services
{
    public class ChatTurn
    {
        public string Role { get; set; } = "user"; // "system" | "user" | "assistant"
        public string Content { get; set; } = string.Empty;
    }

    // Talks to any OpenAI-compatible /chat/completions endpoint (OpenRouter by
    // default, or a custom base URL for another compatible provider).
    public static class AiChatService
    {
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };

        public static async Task<string> SendAsync(AiSettings settings, List<ChatTurn> history)
        {
            var payload = new
            {
                model = settings.Model,
                messages = history.Select(m => new { role = m.Role, content = m.Content }).ToArray()
            };

            using var req = new HttpRequestMessage(HttpMethod.Post, settings.BaseUrl);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
            if (settings.Provider == "OpenRouter")
            {
                req.Headers.Add("HTTP-Referer", "https://miniexel.local");
                req.Headers.Add("X-Title", "Excel Lite + AI");
            }
            req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

            var resp = await Http.SendAsync(req);
            var body = await resp.Content.ReadAsStringAsync();

            if (!resp.IsSuccessStatusCode)
                throw new InvalidOperationException($"Error del proveedor de IA ({(int)resp.StatusCode}): {body}");

            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? string.Empty;
        }

        // Looks for a ```json { ... } ``` block in the assistant's reply so it can
        // propose spreadsheet edits (and new sheets), not just chat. Accepts both
        // the current object schema ({"new_sheets":[...],"edits":[...]}) and a
        // bare edits array for backward compatibility.
        public static AiEditPlan? TryExtractEditPlan(string aiText)
        {
            var match = Regex.Match(aiText, "```json\\s*(\\{.*?\\}|\\[.*?\\])\\s*```", RegexOptions.Singleline);
            if (!match.Success) return null;

            var json = match.Groups[1].Value.TrimStart();
            try
            {
                if (json.StartsWith("["))
                    return new AiEditPlan { Edits = JsonSerializer.Deserialize<List<AiCellEdit>>(json) ?? new() };

                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                return JsonSerializer.Deserialize<AiEditPlan>(json, options);
            }
            catch
            {
                return null;
            }
        }
    }
}
