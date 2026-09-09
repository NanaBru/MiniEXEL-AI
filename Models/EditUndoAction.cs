namespace MiniEXEL.Models
{
    // One reversible cell write, used for Undo/Redo. A single user action (a
    // paste, an AI-applied plan) can produce several of these grouped in a batch.
    public class EditUndoAction
    {
        public string SheetName { get; set; } = string.Empty;
        public int RowNumber { get; set; }
        public int ColNumber { get; set; }
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
    }
}
