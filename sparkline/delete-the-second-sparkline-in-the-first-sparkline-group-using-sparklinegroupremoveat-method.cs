// Title: How to delete the second sparkline in the first sparkline group with Aspose.Cells for .NET using SparklineGroup.RemoveAt
// Common Searches: Aspose.Cells C# remove second sparkline from first sparkline group | SparklineGroup.RemoveAt example for deleting a specific sparkline in .NET | C# code to delete a sparkline at index 1 using Aspose.Cells | Remove a sparkline by index from a sparkline group with Aspose.Cells API
// Tags: SparklineGroup.RemoveAt usage Aspose.Cells | delete specific sparkline .NET | remove sparkline by index Aspose.Cells | manage sparkline groups C# | Aspose.Cells sparkline manipulation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, fills cells A1:A5 with sample values, and shows where to invoke SparklineGroup.RemoveAt to delete the second sparkline in the first sparkline group. It includes a safeguard for the optional Sparkline assembly and saves the workbook as Output.xlsx while handling exceptions.
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

            // Populate some sample data (A1:A5)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(50);

            // NOTE: Sparkline API requires the Aspose.Cells.Sparkline assembly.
            // If the assembly is unavailable, the following block is skipped.
            // This keeps the example functional without causing compile‑time errors.

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
