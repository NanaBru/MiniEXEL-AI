using System.Collections.Generic;

namespace MiniEXEL.Models
{
    public class SheetSnapshot
    {
        public string Name { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public List<ExcelRowVm> Rows { get; set; } = new();
        public List<PictureSnapshot> Pictures { get; set; } = new();
        public bool Truncated { get; set; }
    }
}
