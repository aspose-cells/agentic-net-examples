using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsUtf8CompressionDemo
{
    // Custom LightCellsDataProvider that writes cell values to the target workbook.
    // String values are converted to UTF‑8 byte arrays before being placed into the cell.
    class Utf8CompressProvider : LightCellsDataProvider
    {
        private readonly Worksheet _sourceSheet;
        private int _currentRow = -1;
        private int _currentCol = -1;
        private readonly int _maxRow;
        private readonly int _maxCol;

        public Utf8CompressProvider(Worksheet sourceSheet)
        {
            _sourceSheet = sourceSheet ?? throw new ArgumentNullException(nameof(sourceSheet));
            // Determine the used range of the source sheet.
            var maxRow = sourceSheet.Cells.MaxDataRow;
            var maxCol = sourceSheet.Cells.MaxDataColumn;
            _maxRow = maxRow;
            _maxCol = maxCol;
        }

        // Process only the first worksheet.
        public bool StartSheet(int sheetIndex)
        {
            return sheetIndex == 0;
        }

        // Return the next row index to be saved, or -1 when done.
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1;
            return _currentRow <= _maxRow ? _currentRow : -1;
        }

        // No special row initialization required.
        public void StartRow(Row row) { }

        // Return the next column index within the current row, or -1 when the row ends.
        public int NextCell()
        {
            _currentCol++;
            return _currentCol <= _maxCol ? _currentCol : -1;
        }

        // Write the cell value to the target workbook.
        // If the source cell contains a string, convert it to UTF‑8 bytes first.
        public void StartCell(Cell cell)
        {
            Cell srcCell = _sourceSheet.Cells[_currentRow, _currentCol];

            if (srcCell.Type == CellValueType.IsString)
            {
                string text = srcCell.StringValue;
                byte[] utf8Bytes = Encoding.UTF8.GetBytes(text);
                // Store the byte array directly. Aspose.Cells treats a byte[] as a binary value.
                cell.PutValue(utf8Bytes);
            }
            else
            {
                // Preserve non‑string values as they are.
                cell.PutValue(srcCell.Value);
            }
        }

        // No string gathering needed for this scenario.
        public bool IsGatherString()
        {
            return false;
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Prepare source workbook with sample data ----------
            Workbook sourceWb = new Workbook();
            Worksheet srcWs = sourceWb.Worksheets[0];
            srcWs.Cells["A1"].PutValue("Hello");
            srcWs.Cells["B1"].PutValue("World");
            srcWs.Cells["A2"].PutValue(123);
            srcWs.Cells["B2"].PutValue(45.67);
            srcWs.Cells["A3"].PutValue("Aspose.Cells");
            srcWs.Cells["B3"].PutValue(DateTime.Now);

            // ---------- Save using LightCells mode with UTF‑8 compression ----------
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new Utf8CompressProvider(srcWs),
                // Optional: choose a higher compression level for the package itself.
                CompressionType = OoxmlCompressionType.Level6
            };

            // The target workbook is empty; the provider supplies all cell data.
            Workbook targetWb = new Workbook();
            targetWb.Save("CompressedUtf8.xlsx", saveOptions);

            Console.WriteLine("Workbook saved with UTF‑8 byte array compression.");
        }
    }
}