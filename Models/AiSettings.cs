namespace MiniEXEL.Models
{
    public class AiSettings
    {
        public string Provider { get; set; } = "OpenRouter"; // "OpenRouter" or "Custom"
        public string BaseUrl { get; set; } = "https://openrouter.ai/api/v1/chat/completions";
        public string Model { get; set; } = "openai/gpt-4o-mini";
        public string ApiKey { get; set; } = string.Empty;

        public bool IsConfigured => !string.IsNullOrWhiteSpace(ApiKey) && !string.IsNullOrWhiteSpace(Model) && !string.IsNullOrWhiteSpace(BaseUrl);
    }
}
