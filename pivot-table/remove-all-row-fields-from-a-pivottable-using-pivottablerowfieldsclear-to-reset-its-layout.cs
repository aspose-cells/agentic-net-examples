// Title: How to clear all row fields of a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove every row field from an existing PivotTable | reset PivotTable layout by clearing row fields with Aspose.Cells .NET | example code for PivotTable.RowFields.Clear in C# | programmatically delete all row fields from an Excel PivotTable using Aspose
// Tags: clear pivot table row fields Aspose.Cells | reset pivot table layout Aspose.Cells C# | remove all row fields from Excel PivotTable .NET | Aspose.Cells RowFields.Clear method usage | modify pivot table structure programmatically C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable related classes

// The sample loads an Excel workbook, checks for a PivotTable on the first worksheet, clears every row field with the PivotTable.RowFields.Clear method to reset the layout, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
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
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
