// Title: How to preserve user‑applied cell formatting when refreshing a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: How can I keep the original cell styles when calling PivotTable.RefreshData with Aspose.Cells in C#? | Aspose.Cells C# example to enable PreserveCellFormattingOnUpdate for a PivotTable | Refresh an Excel PivotTable using Aspose.Cells without losing custom formatting | Set PreserveCellFormattingOnUpdate flag in .NET to maintain pivot table formatting after data update | C# code to preserve user‑applied formatting when updating a PivotTable with Aspose.Cells
// Tags: PreserveCellFormattingOnUpdate Aspose.Cells | pivot table update retain formatting .NET | Aspose.Cells maintain pivot styles after refresh | C# workbook pivot table formatting preservation | Excel pivot table formatting flag Aspose

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, accesses the first worksheet and its first PivotTable, enables the PreserveCellFormattingOnUpdate flag to keep user‑applied styles, refreshes the PivotTable data, recalculates formulas, and saves the result to a new file.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet (adjust index if multiple)
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found on the worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Preserve user‑applied formatting after refresh
            pivotTable.PreserveCellFormattingOnUpdate = true;

            // Refresh the PivotTable data source
            try
            {
                pivotTable.RefreshData();
            }
            catch (Exception refreshEx)
            {
                Console.WriteLine($"Warning: Unable to refresh PivotTable data. {refreshEx.Message}");
            }

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
