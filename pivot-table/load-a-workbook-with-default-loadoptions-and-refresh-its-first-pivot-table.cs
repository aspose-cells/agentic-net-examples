// Title: Refresh the first pivot table in an Excel workbook using Aspose.Cells for .NET with default LoadOptions
// Common Searches: c# aspnet refresh pivot table after loading workbook with Aspose.Cells | using Aspose.Cells default LoadOptions to update first pivot table in Excel file | how to programmatically recalculate a pivot table in .NET with Aspose.Cells | Aspose.Cells RefreshData and CalculateData on first worksheet pivot table | load Excel file and refresh its pivot tables using Aspose.Cells for .NET
// Tags: Aspose.Cells RefreshData method | LoadOptions default workbook loading .NET | first worksheet pivot recalculation | Excel pivot table programmatic update | Aspose.Cells pivot table operations

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an Excel file with default LoadOptions, checks the first worksheet for a pivot table, calls RefreshData() and CalculateData() on that pivot, and saves the workbook, demonstrating how to programmatically update a pivot table using Aspose.Cells for .NET.
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook with default LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Refresh the first pivot table if it exists
            if (sheet.PivotTables.Count > 0)
            {
                PivotTable pivot = sheet.PivotTables[0];
                pivot.RefreshData();      // Refresh the data source
                pivot.CalculateData();    // Recalculate the pivot table
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
