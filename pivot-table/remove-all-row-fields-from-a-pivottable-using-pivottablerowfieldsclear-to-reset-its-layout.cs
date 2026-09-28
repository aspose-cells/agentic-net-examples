// Title: How to clear all row fields from a PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, retrieves the first PivotTable, calls RowFields.Clear to remove every row field, and saves the updated file. | Show a .NET console example that resets a PivotTable layout by programmatically clearing its row fields using the Aspose.Cells PivotTable API.
// Common Searches: Aspose.Cells C# clear pivot table row fields example | remove all row fields from a PivotTable programmatically with Aspose.Cells | reset pivot table layout by clearing row fields in .NET | how to use PivotTable.RowFields.Clear in Aspose.Cells | C# code to delete row fields from Excel pivot table using Aspose.Cells
// Tags: Aspose.Cells PivotTable.RowFields.Clear method | C# delete all row fields from Excel pivot | reset pivot layout using Aspose.Cells | remove row fields from pivot table .NET | programmatic clearing of pivot row fields C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads 'input.xlsx', accesses the first worksheet's first PivotTable, clears all its row fields with PivotTable.RowFields.Clear(), and saves the modified workbook as 'output.xlsx', including file existence checks and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Remove all row fields to reset the layout
            pivotTable.RowFields.Clear();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
