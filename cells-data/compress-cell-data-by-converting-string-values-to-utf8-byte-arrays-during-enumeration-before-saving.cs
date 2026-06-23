using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsUtf8CompressionDemo
{
    // Custom LightCellsDataProvider that converts each string value to a UTF‑8 byte array
    // before the cell is written to the output file.
    public class Utf8CompressDataProvider : LightCellsDataProvider
    {
        // Sample data to be written – can be replaced with any source.
        private readonly string[,] _data = new string[,]
        {
            { "ID", "Name", "Description" },
            { "1", "Apple",  "Fresh red apple" },
            { "2", "Banana", "Ripe yellow banana" },
            { "3", "Cherry", "Sweet dark cherry" }
        };

        private int _currentRow = -1;
        private int _currentCol = -1;

        // Process only the first worksheet.
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index or -1 when finished.
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1;
            return _currentRow < _data.GetLength(0) ? _currentRow : -1;
        }

        // No special row handling required.
        public void StartRow(Row row) { }

        // Return the next column index or -1 when the row is finished.
        public int NextCell()
        {
            _currentCol++;
            return _currentCol < _data.GetLength(1) ? _currentCol : -1;
        }

        // Convert the string value to UTF‑8 bytes and store the byte array in the cell.
        public void StartCell(Cell cell)
        {
            string originalValue = _data[_currentRow, _currentCol];

            // Convert the string to a UTF‑8 encoded byte array.
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(originalValue);

            // Store the byte array. Aspose.Cells will treat the byte[] as a binary value.
            cell.PutValue(utf8Bytes);
        }

        // No string gathering is required for this example.
        public bool IsGatherString() => false;
    }

    public class Program
    {
        public static void Main()
        {
            // Create an empty workbook – the actual data will be supplied by the provider.
            Workbook workbook = new Workbook();

            // Configure OoxmlSaveOptions:
            //   * Use the custom LightCellsDataProvider to supply cell data.
            //   * Set the highest compression level (Level9) for maximum size reduction.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new Utf8CompressDataProvider(),
                CompressionType = OoxmlCompressionType.Level9
            };

            // Save the workbook. The provider streams the UTF‑8 byte arrays directly to the file.
            workbook.Save("Utf8CompressedWorkbook.xlsx", saveOptions);

            Console.WriteLine("Workbook saved with UTF‑8 byte array compression and Level9 Ooxml compression.");
        }
    }
}