using System;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
using MiniEXEL.Models;

namespace MiniEXEL.Services
{
    // Full-fidelity reader used by the single-workbook "Office-like" view:
    // colors, fonts, formulas (with cached results) and embedded pictures.
    // Only .xlsx/.xlsm are supported (ClosedXML can't read legacy .xls);
    // for .xls we fall back to a plain-text view via ExcelDataReader.
    public class OpenWorkbookResult
    {
        public bool SupportsRichView { get; set; }
        public XLWorkbook? Workbook { get; set; }
        public System.Collections.Generic.List<string> SheetNames { get; set; } = new();
        public string? LimitationNote { get; set; }
    }

    public static class WorkbookService
    {
        public const int MaxRows = 150;
        public const int MaxCols = 40;

        public static OpenWorkbookResult Open(string path)
        {
            var ext = Path.GetExtension(path);
            if (string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                return new OpenWorkbookResult
                {
                    SupportsRichView = false,
                    SheetNames = new() { "Hoja1" },
                    LimitationNote = "Formato .xls (antiguo): se muestra solo texto, sin colores, fórmulas ni imágenes."
                };
            }

            var wb = new XLWorkbook(path);
            return new OpenWorkbookResult
            {
                SupportsRichView = true,
                Workbook = wb,
                SheetNames = wb.Worksheets.Select(w => w.Name).ToList()
            };
        }

        public static SheetSnapshot BuildSnapshot(XLWorkbook wb, string sheetName)
        {
            var ws = wb.Worksheet(sheetName);
            var used = ws.RangeUsed();

            int rowCount = used == null ? 0 : Math.Min(used.RowCount(), MaxRows);
            int colCount = used == null ? 0 : Math.Min(used.ColumnCount(), MaxCols);
            bool truncated = used != null && (used.RowCount() > MaxRows || used.ColumnCount() > MaxCols);

            // Pictures can be anchored past the used cell range (e.g. a logo placed
            // to the right of the data) — widen the grid so they still render.
            foreach (var pic in ws.Pictures)
            {
                rowCount = Math.Max(rowCount, Math.Min(pic.TopLeftCell.Address.RowNumber, MaxRows));
                colCount = Math.Max(colCount, Math.Min(pic.TopLeftCell.Address.ColumnNumber, MaxCols));
            }

            var rows = new System.Collections.Generic.List<ExcelRowVm>(rowCount);

            for (int r = 1; r <= rowCount; r++)
            {
                var cells = new CellSnapshot[colCount];
                for (int c = 1; c <= colCount; c++)
                {
                    var cell = ws.Cell(r, c);

                    string display;
                    try
                    {
                        // cell.Value evaluates formulas on demand when the file has no
                        // cached result (e.g. produced by a tool other than Excel).
                        display = cell.Value.ToString() ?? string.Empty;
                    }
                    catch
                    {
                        // Formula engine couldn't evaluate this cell (unsupported function, etc).
                        display = cell.HasFormula ? "#ERROR" : cell.GetString();
                    }

                    var snapshot = new CellSnapshot
                    {
                        Address = cell.Address.ToString(),
                        RowNumber = r,
                        ColNumber = c,
                        DisplayValue = display,
                        Formula = cell.HasFormula ? cell.FormulaA1 : null,
                        Background = ToBrush(cell.Style.Fill.BackgroundColor, Colors.Transparent),
                        Foreground = ToBrush(cell.Style.Font.FontColor, Colors.Black),
                        FontWeightValue = cell.Style.Font.Bold ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal
                    };
                    snapshot.Background.Freeze();
                    snapshot.Foreground.Freeze();

                    cells[c - 1] = snapshot;
                }
                rows.Add(new ExcelRowVm { RowNumber = r, Cells = cells });
            }

            var pictures = new System.Collections.Generic.List<PictureSnapshot>();
            foreach (var pic in ws.Pictures)
            {
                try
                {
                    int anchorRow = pic.TopLeftCell.Address.RowNumber - 1;
                    int anchorCol = pic.TopLeftCell.Address.ColumnNumber - 1;
                    if (anchorRow >= rowCount || anchorCol >= colCount) continue; // outside the bounded preview area

                    using var src = pic.ImageStream;
                    using var ms = new MemoryStream();
                    src.Position = 0;
                    src.CopyTo(ms);
                    ms.Position = 0;

                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.StreamSource = ms;
                    bmp.EndInit();
                    bmp.Freeze();

                    pictures.Add(new PictureSnapshot
                    {
                        Image = bmp,
                        AnchorRowIndex = anchorRow,
                        AnchorColIndex = anchorCol,
                        WidthPx = pic.Width,
                        HeightPx = pic.Height
                    });
                }
                catch
                {
                    // Skip pictures in formats we can't decode; the rest of the sheet still renders.
                }
            }

            return new SheetSnapshot
            {
                Name = sheetName,
                RowCount = rowCount,
                ColCount = colCount,
                Rows = rows,
                Pictures = pictures,
                Truncated = truncated
            };
        }

        // Writes a cell edited via the grid or the AI assistant back into the live
        // workbook (kept open in memory while the WorkbookView is active).
        public static void ApplyEdit(XLWorkbook wb, string sheetName, int row, int col, string rawText)
        {
            var ws = wb.Worksheet(sheetName);
            var cell = ws.Cell(row, col);

            if (rawText.StartsWith("="))
                cell.FormulaA1 = rawText.Substring(1);
            else
                cell.Value = rawText;
        }

        // Creates a sheet if it doesn't already exist (used by the AI assistant
        // when it proposes a "new_sheets" plan). Returns true if a sheet was created.
        public static bool EnsureSheetExists(XLWorkbook wb, string sheetName)
        {
            if (wb.Worksheets.Contains(sheetName)) return false;
            wb.AddWorksheet(sheetName);
            return true;
        }

        public static bool TryApplyEditByAddress(XLWorkbook wb, string sheetName, string address, string rawText)
        {
            try
            {
                var ws = wb.Worksheet(sheetName);
                var cell = ws.Cell(address);
                if (rawText.StartsWith("="))
                    cell.FormulaA1 = rawText.Substring(1);
                else
                    cell.Value = rawText;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static SolidColorBrush ToBrush(XLColor xlColor, Color fallback)
        {
            try
            {
                if (xlColor.ColorType == XLColorType.Color)
                {
                    var c = xlColor.Color;
                    return new SolidColorBrush(Color.FromArgb(c.A, c.R, c.G, c.B));
                }
            }
            catch
            {
                // Themed/indexed colors ClosedXML can't resolve fall back below.
            }
            return new SolidColorBrush(fallback);
        }
    }
}
