using System;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace MiniEXEL.Services
{
    public class SheetPreview
    {
        public string SheetName { get; set; } = string.Empty;
        public DataTable Table { get; set; } = new DataTable();
        public bool Truncated { get; set; }
    }

    // Reads only a bounded number of rows/columns from the first sheet using
    // ExcelDataReader's forward-only streaming reader, so previewing a huge
    // workbook never loads the whole file into memory.
    public static class ExcelPreviewService
    {
        private const int MaxRows = 200;
        private const int MaxCols = 60;

        public static SheetPreview LoadPreview(string filePath)
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var table = new DataTable();
            bool truncated = false;
            string sheetName = reader.Name ?? "Sheet1";

            bool headerRead = false;
            int rowCount = 0;

            do
            {
                if (!string.Equals(reader.Name, sheetName, StringComparison.Ordinal))
                    break; // only preview the first sheet

                while (reader.Read())
                {
                    int fieldCount = Math.Min(reader.FieldCount, MaxCols);

                    if (!headerRead)
                    {
                        for (int c = 0; c < fieldCount; c++)
                        {
                            var colName = reader.GetValue(c)?.ToString();
                            table.Columns.Add(string.IsNullOrWhiteSpace(colName) ? $"Col{c + 1}" : colName);
                        }
                        if (reader.FieldCount > MaxCols) truncated = true;
                        headerRead = true;
                        continue;
                    }

                    if (rowCount >= MaxRows)
                    {
                        truncated = true;
                        break;
                    }

                    var row = table.NewRow();
                    for (int c = 0; c < fieldCount; c++)
                        row[c] = reader.GetValue(c) ?? DBNull.Value;
                    table.Rows.Add(row);
                    rowCount++;
                }
                break;
            } while (reader.NextResult());

            return new SheetPreview { SheetName = sheetName, Table = table, Truncated = truncated };
        }
    }
}
