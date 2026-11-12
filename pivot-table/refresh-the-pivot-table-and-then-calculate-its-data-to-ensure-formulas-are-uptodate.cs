// Title: How to refresh a PivotTable, calculate its data, and update workbook formulas using Aspose.Cells in C#
// Common Searches: C# code to refresh an Excel PivotTable and recalculate data with Aspose.Cells | Aspose.Cells .NET example for refreshing pivot source and updating formulas | Programmatically recalculate workbook formulas after pivot refresh using Aspose.Cells | How to call RefreshData and CalculateData on a PivotTable in Aspose.Cells C# | Update PivotTable source data and recalculate workbook in Aspose.Cells .NET
// Tags: Aspose.Cells pivot table refresh | PivotTable CalculateData method | Workbook.CalculateFormula usage | C# update Excel pivot source | Aspose.Cells data recalculation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel workbook, checks for a PivotTable on the first worksheet, refreshes its source data, calculates the PivotTable values, recalculates all workbook formulas, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh the pivot table's source data and recalculate its values
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Recalculate all workbook formulas to keep them up‑to‑date
            workbook.CalculateFormula();

            // Save the modified workbook
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
