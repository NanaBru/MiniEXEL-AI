using System.Collections.Generic;

namespace MiniEXEL.Models
{
    // One proposed change from the AI assistant, e.g. {"cell":"D2","value":"=B2*C2"}.
    // "Sheet" is optional — omitted means "the sheet currently open in that pane";
    // set it (matching a name in NewSheets, or an existing sheet) to target another one.
    public class AiCellEdit
    {
        public string? Sheet { get; set; }
        public string Cell { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    // Full plan the AI can propose in one reply: create sheets and/or edit cells.
    public class AiEditPlan
    {
        public List<string> NewSheets { get; set; } = new();
        public List<AiCellEdit> Edits { get; set; } = new();
    }
}
