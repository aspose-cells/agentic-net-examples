using Aspose.Cells;
using System;

class ImportArrayDemo
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data to import – an object array containing mixed types
        object[] data = new object[]
        {
            "Name", "Age", "City",          // Header row
            "Alice", 30, "New York",        // First data row
            "Bob",   25, "Los Angeles",    // Second data row
            "Carol", 28, "Chicago"         // Third data row
        };

        // Import the array horizontally starting at cell A1 (row 0, column 0)
        // isVertical = false means data will be placed left‑to‑right, then wrap to next row
        worksheet.Cells.ImportObjectArray(data, 0, 0, false);

        // Save the workbook in XLSX format
        workbook.Save("ArrayImportDemo.xlsx");
    }
}