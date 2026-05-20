using System;
using System.IO;
using Aspose.Cells;

class ClearMappedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in the range A1:B2
            cells["A1"].PutValue("John");
            cells["B1"].PutValue(30);
            cells["A2"].PutValue("Mary");
            cells["B2"].PutValue(25);

            // Define the CellArea that corresponds to the range to be cleared
            CellArea mappedArea = new CellArea
            {
                StartRow = 0,      // Row 0 (A)
                StartColumn = 0,   // Column 0 (1)
                EndRow = 1,        // Row 1 (2)
                EndColumn = 1      // Column 1 (B)
            };

            // Clear only the contents of the defined range
            cells.ClearContents(mappedArea);

            // Save the workbook
            string outputPath = "ClearedMappedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}