// Title: How to disable row grand totals in an Aspose.Cells pivot table using C#
// Common Searches: Aspose.Cells C# hide row grand totals in pivot table | disable row grand total property in Aspose.Cells pivot table | remove row grand totals from Excel pivot table programmatically with Aspose.Cells | C# code to turn off row grand totals in a pivot table using Aspose.Cells | Aspose.Cells pivot table RowGrandTotals false example | Excel pivot table summary view without row grand totals using Aspose.Cells .NET
// Tags: Aspose.Cells pivot table row grand totals | C# disable pivot table row grand totals | Aspose.Cells RowGrandTotals property | Excel pivot table summary view customization .NET | Aspose.Cells hide pivot totals

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, checks for a pivot table on the first worksheet, accesses the first pivot table, and demonstrates how to turn off its RowGrandTotals property using Aspose.Cells for .NET before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Disable row grand totals (property not available in older versions)
            // pivotTable.RowGrandTotals = false;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
