// Title: C# example: Refresh a PivotTable, calculate its data, and recalculate workbook formulas with Aspose.Cells
// Common Searches: how to refresh a pivot table and recalculate data using Aspose.Cells in C# | Aspose.Cells C# code to update pivot source and run CalculateData | refresh pivot table and recalculate workbook formulas Aspose.Cells .NET | example of PivotTable.RefreshData and CalculateData with Aspose.Cells | C# program to load Excel, refresh pivot, and recalculate formulas using Aspose
// Tags: Aspose.Cells PivotTable.RefreshData C# | Aspose.Cells PivotTable.CalculateData example | recalculate workbook formulas Aspose.Cells .NET | refresh pivot source data Aspose.Cells | C# load workbook update pivot Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The C# program loads an existing Excel workbook, verifies a pivot table on the first worksheet, refreshes its source data, calculates the pivot data, recalculates all workbook formulas, and saves the updated file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
                throw new InvalidOperationException("No pivot tables found in the worksheet.");

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh source data and recalculate pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Recalculate all workbook formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
