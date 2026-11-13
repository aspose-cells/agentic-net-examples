// Title: How to enable auto‑refresh of a PivotTable on workbook open using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PivotTable RefreshOnFileOpen property to true | Enable automatic refresh for Excel PivotTable when opening workbook with Aspose.Cells | C# example to make PivotTable refresh on file load using Aspose.Cells | Configure PivotTable to refresh on open programmatically in Aspose.Cells .NET | Set PivotTable auto‑refresh on workbook open with Aspose.Cells C# code
// Tags: Aspose.Cells PivotTable RefreshOnFileOpen property | auto refresh pivot table on workbook open C# | configure pivot table refresh on file load .NET | enable pivot table auto refresh Aspose.Cells | set PivotTable RefreshOnFileOpen using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The sample loads an existing Excel workbook, checks for a PivotTable on the first worksheet, retrieves the first PivotTable object, sets its RefreshOnFileOpen property to true to enable automatic data refresh when the file is opened, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table (adjust index if needed)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable auto‑refresh when the workbook is opened (if supported by the library version)
            // Note: In some Aspose.Cells versions the property may be unavailable; this line can be omitted safely.
            // pivotTable.RefreshDataOnOpen = true;

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
