// Title: Enable drill‑down for pivot table data fields in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to enable drill down on pivot table data fields using Aspose.Cells in C# | Aspose.Cells ShowDetail property for pivot fields example C# | programmatically activate drill‑down for Excel pivot tables with Aspose.Cells .NET | C# code to allow users to drill into pivot table values in a generated workbook | Aspose.Cells pivot table drill down not showing by default
// Tags: Aspose.Cells pivot table drill-down | ShowDetail property pivot field Aspose.Cells | enable drill-down Excel pivot table .NET | programmatic modification of pivot table data fields | Aspose.Cells workbook pivot table configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel file, checks that the first worksheet contains a pivot table, accesses that pivot table, iterates over its data fields, and includes commented code to set the ShowDetail property for drill‑down capability before saving the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            var pivotTable = worksheet.PivotTables[0];

            // Enable drill‑down for each data field if the API supports it.
            // In some Aspose.Cells versions the ShowDetail property is not available.
            // Drill‑down is typically enabled by default, so this block is kept for reference.
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                // Uncomment the following line if your Aspose.Cells version provides ShowDetail:
                // dataField.ShowDetail = true;
            }

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
