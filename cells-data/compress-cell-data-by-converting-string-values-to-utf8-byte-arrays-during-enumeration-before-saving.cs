// Title: Encode string cells to Base64 UTF‑8 during LightCells enumeration and save XLSX with OOXML Level 9 compression using Aspose.Cells for .NET
// AI Prompts: Create a LightCellsDataProvider that checks each cell value, converts string values to UTF‑8 byte arrays, encodes those bytes to Base64, and writes the encoded string back to the cell in the StartCell method. | Set up OoxmlSaveOptions to use the custom LightCellsDataProvider and specify CompressionType = Level9, then save the workbook as an XLSX file.
// Common Searches: how to use LightCellsDataProvider to compress string values in Aspose.Cells | save Excel workbook with maximum OOXML compression in C# Aspose.Cells | convert cell text to Base64 during streaming export with Aspose.Cells | custom data provider for on‑the‑fly string encoding in Aspose.Cells .NET
// Tags: LightCellsDataProvider string Base64 encoding | OOXML Level9 compression Aspose.Cells | UTF8 byte array conversion C# | streaming workbook save Aspose.Cells | custom cell value compression .NET

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCompressionDemo
{
    // Custom LightCellsDataProvider that converts string values to UTF‑8 byte arrays
    // during enumeration (StartCell) before the workbook is saved.
    // Demonstrates a LightCellsDataProvider that transforms string cell values into UTF‑8 byte arrays, encodes them as Base64, and writes them during enumeration, while enabling string gathering and saving the workbook with OOXML Level 9 compression for optimal file size.
    public class Utf8CompressingDataProvider : LightCellsDataProvider
    {
        private int _currentRow = -1;
        private int _currentCol = -1;

        // Only process the first worksheet.
        public bool StartSheet(int sheetIndex) => sheetIndex == 0;

        // Return the next row index to be saved.
        public int NextRow()
        {
            _currentRow++;
            _currentCol = -1;
            // Assume we have data for rows 0‑9.
            return _currentRow < 10 ? _currentRow : -1;
        }

        // No special row handling needed.
        public void StartRow(Row row) { }

        // Return the next column index to be saved.
        public int NextCell()
        {
            _currentCol++;
            // Assume we have data for columns 0‑4.
            return _currentCol < 5 ? _currentCol : -1;
        }

        // Fill the cell value. If the value is a string, convert it to a UTF‑8 byte array
        // and then back to a string (simulating compression before saving).
        public void StartCell(Cell cell)
        {
            // Example data source: a simple 2‑dimensional array.
            // In a real scenario this could be any external source.
            object[,] sourceData = new object[,]
            {
                { "ID", "Name", "Description", "Price", "Notes" },
                { 1, "Item A", "A long description that could be compressed.", 12.5, "Special" },
                { 2, "Item B", "Another description.", 8.75, "Standard" },
                { 3, "Item C", "Short desc.", 15.0, "Premium" },
                { 4, "Item D", "Description with repeated words words words.", 9.99, "Discount" },
                { 5, "Item E", "Final item description.", 20.0, "New" },
                { 6, "Item F", "Extra info.", 7.5, "Clearance" },
                { 7, "Item G", "More details.", 13.3, "Limited" },
                { 8, "Item H", "Sample text.", 11.1, "Regular" },
                { 9, "Item I", "End of list.", 5.0, "Last" }
            };

            object value = sourceData[_currentRow, _currentCol];

            // If the value is a string, compress it by converting to UTF‑8 bytes.
            if (value is string str)
            {
                // Convert the string to UTF‑8 bytes.
                byte[] utf8Bytes = Encoding.UTF8.GetBytes(str);

                // Here we could store the byte array in a custom binary format.
                // For demonstration, we convert the bytes back to a Base64 string,
                // which reduces the size when many duplicate strings exist.
                string compressed = Convert.ToBase64String(utf8Bytes);

                cell.PutValue(compressed);
            }
            else
            {
                // Non‑string values are written unchanged.
                cell.PutValue(value);
            }
        }

        // Enable string gathering to further reduce duplication in the final file.
        public bool IsGatherString() => true;
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty – data will be supplied by the provider).
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear(); // Remove default sheet; provider will create rows/cells.

            // Configure save options:
            // - Use the custom LightCellsDataProvider.
            // - Apply the highest compression level for the OOXML package.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new Utf8CompressingDataProvider(),
                CompressionType = OoxmlCompressionType.Level9   // Best compression.
            };

            // Save the workbook. The provider streams data directly to the file,
            // applying the UTF‑8 conversion and compression on the fly.
            workbook.Save("CompressedData.xlsx", saveOptions);

            Console.WriteLine("Workbook saved with UTF‑8 string compression and OOXML Level9 compression.");
        }
    }
}
