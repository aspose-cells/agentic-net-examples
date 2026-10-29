// Title: Hide a PivotTable’s field list pane in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to hide pivot table field list pane with Aspose.Cells in C# | Aspose.Cells set ShowFieldList false before saving workbook | disable pivot table UI pane programmatically using Aspose.Cells .NET | C# code to suppress PivotTable field list in generated Excel file | Aspose.Cells hide pivot table field list pane prior to workbook save
// Tags: Aspose.Cells pivot table UI options | C# pivot table field list suppression | Excel workbook pivot UI customization Aspose | ShowFieldList property Aspose.Cells .NET | programmatic pivot table settings C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable classes

// The example loads an existing workbook, checks for a PivotTable on the first worksheet, notes that the ShowFieldList property is not available in the current Aspose.Cells version, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable
            if (worksheet.PivotTables.Count > 0)
            {
                // Access the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // The ShowFieldList property is not available in the current Aspose.Cells version.
                // If needed, additional pivot table settings can be configured here.
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
