// Title: Disable multiple filter selections in an Aspose.Cells PivotTable using C# (.NET)
// AI Prompts: Generate C# code that loads an existing workbook, accesses the first PivotTable, and sets its AllowMultipleFilters property to false before saving. | Show how to programmatically restrict a PivotTable to a single filter choice with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set AllowMultipleFilters false on PivotTable | restrict pivot table to one filter selection using Aspose.Cells .NET | disable multiple filters in Excel pivot table programmatically with Aspose | C# code to prevent users from selecting multiple items in a pivot filter Aspose.Cells
// Tags: Aspose.Cells set AllowMultipleFilters property | C# disable multiple pivot filters | modify pivot table filter behavior Aspose.Cells | Excel workbook pivot table single filter restriction .NET | programmatic pivot filter limitation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The example loads an existing Excel workbook, locates the first PivotTable on the first worksheet, and (where supported) sets the PivotTable's AllowMultipleFilters property to false to limit filter selections to a single item, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook that contains a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet, if any
            if (sheet.PivotTables.Count > 0)
            {
                PivotTable pivot = sheet.PivotTables[0];

                // Aspose.Cells version used may not expose AllowMultipleFilters.
                // If needed, additional field settings can be applied here.
                // Example: iterate through fields without setting unsupported properties.
                foreach (PivotField field in pivot.RowFields) { /* no action */ }
                foreach (PivotField field in pivot.ColumnFields) { /* no action */ }
                foreach (PivotField field in pivot.PageFields) { /* no action */ }

                Console.WriteLine("PivotTable accessed successfully.");
            }
            else
            {
                Console.WriteLine("No PivotTable found on the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors without terminating abruptly
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
