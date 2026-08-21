// Title: Compress Excel strings with UTF‑8 using a custom LightCellsDataProvider in Aspose.Cells for .NET
// Description: This example creates a workbook, fills it with mixed data, and saves it as an XLSX file using OoxmlSaveOptions with level‑9 compression. A custom LightCellsDataProvider (Utf8CompressProvider) enumerates the first worksheet, converts each string cell to a UTF‑8 byte array (and back to a string) during export, copies non‑string values unchanged, and enables duplicate‑string gathering to further reduce file size.
// Keywords: Aspose.Cells | LightCellsDataProvider | UTF-8 string compression | OoxmlSaveOptions | Level9 compression | duplicate string gathering | C# | .NET | Excel export optimization | custom provider
// Common Searches: How to compress string data in Aspose.Cells using LightCellsDataProvider | Save workbook with maximum OOXML compression in .NET | Enable duplicate string gathering when exporting XLSX with Aspose.Cells | Convert Excel cell strings to UTF‑8 bytes during save
// Developer Intent: Export a workbook with maximum OOXML compression while converting string cells to UTF‑8 byte arrays via a custom LightCellsDataProvider.
// Use Cases: Minimize XLSX size by gathering duplicate strings and applying level‑9 compression. | Pre‑process or encode string cells (e.g., UTF‑8, sanitization) without affecting numeric data. | Handle large worksheets with low memory overhead by streaming cell values through a custom provider.
// AI Prompts: Generate a LightCellsDataProvider for Aspose.Cells that encrypts string values before saving. | Modify the Utf8CompressProvider to skip empty cells and only process non‑empty strings. | Explain the impact of IsGatherString() on file size when using OoxmlSaveOptions with level‑9 compression.

using System;
using System.Text;
using Aspose.Cells;

// This example creates a workbook, fills it with mixed data, and saves it as an XLSX file using OoxmlSaveOptions with level‑9 compression. A custom LightCellsDataProvider (Utf8CompressProvider) enumerates the first worksheet, converts each string cell to a UTF‑8 byte array (and back to a string) during export, copies non‑string values unchanged, and enables duplicate‑string gathering to further reduce file size.
class Program
{
    static void Main()
    {
        // Create a workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["A2"].PutValue("World");
        sheet.Cells["B1"].PutValue("Sample");
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["C1"].PutValue("Hello"); // duplicate string to benefit from gathering

        // Configure save options: highest compression and custom LightCellsDataProvider
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            CompressionType = OoxmlCompressionType.Level9, // maximum compression
            LightCellsDataProvider = new Utf8CompressProvider(workbook) // custom provider
        };

        // Save the workbook using the configured options
        workbook.Save("CompressedUtf8.xlsx", saveOptions);
    }
}

// Custom LightCellsDataProvider that converts string values to UTF‑8 byte arrays
// during enumeration before the workbook is saved.
class Utf8CompressProvider : LightCellsDataProvider
{
    private readonly Workbook _sourceWorkbook;
    private int _currentRow = -1;
    private int _currentCol = -1;
    private readonly int _maxRow;
    private readonly int _maxCol;

    public Utf8CompressProvider(Workbook sourceWorkbook)
    {
        _sourceWorkbook = sourceWorkbook;

        // Determine the used range of the first worksheet
        Worksheet ws = _sourceWorkbook.Worksheets[0];
        _maxRow = ws.Cells.MaxDataRow;
        _maxCol = ws.Cells.MaxDataColumn;
    }

    // Process only the first worksheet
    public bool StartSheet(int sheetIndex) => sheetIndex == 0;

    // Return the next row index or -1 when done
    public int NextRow()
    {
        _currentRow++;
        _currentCol = -1;
        return _currentRow <= _maxRow ? _currentRow : -1;
    }

    // No special row initialization required
    public void StartRow(Row row) { }

    // Return the next column index or -1 when done
    public int NextCell()
    {
        _currentCol++;
        return _currentCol <= _maxCol ? _currentCol : -1;
    }

    // Populate the cell being saved
    public void StartCell(Cell cell)
    {
        // Get the original cell from the source workbook
        Cell src = _sourceWorkbook.Worksheets[0].Cells[_currentRow, _currentCol];

        if (src.Type == CellValueType.IsString)
        {
            // Convert the string to UTF‑8 bytes and back to string.
            // This simulates processing the string data before saving.
            string original = src.StringValue;
            byte[] utf8Bytes = Encoding.UTF8.GetBytes(original);
            string processed = Encoding.UTF8.GetString(utf8Bytes);
            cell.PutValue(processed);
        }
        else
        {
            // Copy non‑string values unchanged
            cell.PutValue(src.Value);
        }
    }

    // Enable gathering of duplicate strings to reduce file size
    public bool IsGatherString() => true;
}
