// Title: Enable ImportTableOptions.CheckMergedCells to keep merged cells while importing custom objects with Aspose.Cells for .NET
// AI Prompts: Write C# code that imports an array of custom objects into a worksheet using Aspose.Cells, setting the CheckMergedCells flag on ImportTableOptions together with field name headers. | Show how to configure ImportTableOptions to preserve merged cell layouts before calling Cells.ImportCustomObjects in a .NET project. | Provide a complete example that creates a workbook, adds merged cells, and imports custom objects without losing the merged ranges.
// Common Searches: Aspose.Cells ImportCustomObjects keep merged cells | Set CheckMergedCells true when importing data in C# | Import custom objects into Excel worksheet with merged cells using Aspose.Cells | How to avoid losing merged cell formatting with ImportTableOptions | C# example for ImportTableOptions.CheckMergedCells property
// Tags: ImportTableOptions.CheckMergedCells | ImportCustomObjects merged cell preservation | Aspose.Cells custom object import with headers | C# Excel workbook merged cells handling | ImportTableOptions IsFieldNameShown usage

using System;
using Aspose.Cells;

// The sample creates a workbook, defines custom objects, and imports them with ImportTableOptions that show field names. To correctly populate merged cell ranges, the CheckMergedCells property must be set to true before calling Cells.ImportCustomObjects.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Sample custom objects to import
            var customData = new[]
            {
                new { Name = "Alice", Score = 85 },
                new { Name = "Bob",   Score = 92 }
            };

            // Set import options to treat the first row as headers
            var importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true
            };

            // Import the custom objects starting at cell A1
            worksheet.Cells.ImportCustomObjects(customData, 0, 0, importOptions);

            // Save the workbook
            workbook.Save("MergedCellsExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
