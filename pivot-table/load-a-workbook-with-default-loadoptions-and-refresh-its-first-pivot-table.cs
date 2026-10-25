// Title: Refresh the first pivot table in an Excel workbook using Aspose.Cells for .NET with default LoadOptions
// Common Searches: c# load Excel workbook with Aspose.Cells default LoadOptions and recalculate pivot table | how to programmatically update pivot table data in .NET using Aspose.Cells | Aspose.Cells example for refreshing first worksheet pivot after opening file | using Aspose.Cells to recalc pivot table before saving workbook in C# | load workbook and refresh pivot tables with Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table refresh operation | default LoadOptions workbook loading C# | recalculate pivot data Aspose.Cells .NET | first worksheet pivot table update | Excel file pivot table refresh using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

// The code loads an Excel file with default LoadOptions, checks the first worksheet for a pivot table, refreshes its source data and recalculates the pivot, then saves the modified workbook to a new file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook with default LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Refresh the first pivot table if it exists
            if (worksheet.PivotTables.Count > 0)
            {
                PivotTable pivot = worksheet.PivotTables[0];
                pivot.RefreshData();      // Refresh source data
                pivot.CalculateData();    // Recalculate the pivot table
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
