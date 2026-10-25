// Title: How to keep user‑applied formatting when refreshing an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# preserve pivot table formatting after RefreshData | set PreserveFormatting property on PivotTable before RefreshData in .NET | refresh Excel pivot table without losing custom styles using Aspose.Cells | C# example to enable PreserveFormatting for PivotTable and recalculate data | Aspose.Cells keep user formatting when updating pivot cache
// Tags: Aspose.Cells PreserveFormatting for PivotTable | C# refresh PivotTable data source | maintain custom pivot formatting Aspose.Cells | PivotTable CalculateData usage | load workbook modify pivot settings .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an Excel workbook, accesses the first PivotTable, enables the PreserveFormatting flag to retain user‑applied styles, refreshes and recalculates the pivot data, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet that holds the PivotTable (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable PreserveFormatting so that user‑applied styles are retained after refresh
            pivotTable.PreserveFormatting = true;

            // Refresh the PivotTable data from its source
            pivotTable.RefreshData();

            // Recalculate the PivotTable (optional but ensures data is up‑to‑date)
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
