// Title: C# example: Save an Excel workbook with LightCells API while disabling automatic calculation for faster static sheet creation using Aspose.Cells
// Common Searches: how to save an Excel file with LightCells API in C# without triggering formula recalculation | disable automatic calculation when writing static data with Aspose.Cells LightCells in .NET | Aspose.Cells LightCells performance tip: turn off calc mode before saving workbook | C# code to create a workbook, set manual calculation mode, and save using LightCells | fast Excel generation using LightCells API and manual calculation mode in Aspose.Cells
// Tags: LightCells API workbook saving C# | manual calc mode workaround Aspose.Cells | static sheet performance Aspose.Cells .NET | Excel file output without recalculation | optimize LightCells write speed

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, optionally switches the calculation mode to manual, writes a header and a data row, ensures the target directory exists, and saves the file as output.xlsx using the LightCells API, demonstrating how to avoid automatic recalculation for faster static sheet generation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Disable automatic calculation if supported (optional)
            // workbook.Settings.CalcMode = CalcMode.Manual; // Uncomment if CalcMode is available in your Aspose.Cells version

            // Get reference to the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Write header row
            sheet.Cells[0, 0].PutValue("Header1");
            sheet.Cells[0, 1].PutValue("Header2");

            // Write a data row
            sheet.Cells[1, 0].PutValue(123);
            sheet.Cells[1, 1].PutValue(456);

            // Define output path
            string outputPath = "output.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
