namespace MiniEXEL.Models
{
    public class ExcelRowVm
    {
        public int RowNumber { get; set; }
        public CellSnapshot[] Cells { get; set; } = System.Array.Empty<CellSnapshot>();
    }
}
