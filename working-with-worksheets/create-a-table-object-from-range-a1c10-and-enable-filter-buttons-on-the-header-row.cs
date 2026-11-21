// Title: Add a filtered Excel table from range A1:C10 using Aspose.Cells for .NET
// Common Searches: how to create an Excel table with auto‑filter using Aspose.Cells C# | Aspose.Cells ListObject add range A1:C10 and enable header filters | C# example to apply a built‑in table style and filter buttons in an Excel workbook | saving a workbook that contains a filtered table with Aspose.Cells .NET | code sample for showing header row filters on a ListObject in Aspose.Cells
// Tags: Aspose.Cells ListObject creation from range | enable auto‑filter on Excel table header Aspose.Cells | apply built‑in table style Aspose.Cells | save workbook with filtered table .NET | populate sample data A1:C10 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, fills cells A1:C10 with sample data, adds a ListObject covering that range, makes the header row visible to provide filter buttons, applies a built‑in table style, and saves the file as TableWithFilters.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data in the range A1:C10 (optional, for demonstration)
            for (int row = 0; row < 10; row++)
            {
                sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                sheet.Cells[row, 1].PutValue(row * 10);
                sheet.Cells[row, 2].PutValue(DateTime.Today.AddDays(row));
            }

            // Create a table (ListObject) from the range A1:C10
            // Add returns the index of the created ListObject
            int tableIndex = sheet.ListObjects.Add(0, 0, 9, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is visible
            table.ShowHeaderRow = true;

            // Apply a built‑in table style
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Define output file path
            string outputPath = "TableWithFilters.xlsx";

            // Ensure the directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
