// Title: Refresh an Excel PivotTable with PivotTable.CalculateData() before saving using Aspose.Cells for .NET (C#)
// Common Searches: how to refresh pivot table data programmatically with Aspose.Cells in C# | Aspose.Cells PivotTable.CalculateData example for .NET | recalculate Excel pivot cache before saving workbook using Aspose.Cells | C# code to update pivot table after modifying source data with Aspose.Cells | ensure pivot table reflects latest data when saving workbook with Aspose.Cells
// Tags: PivotTable.CalculateData method Aspose.Cells | refresh Excel pivot cache using .NET | recalculate pivot data before saving workbook | Aspose.Cells programmatic pivot table update | C# Excel pivot table refresh operation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// Loads an existing Excel workbook, checks for a PivotTable on the first worksheet, calls PivotTable.CalculateData() to refresh its data, and then saves the workbook, ensuring the pivot reflects the latest values.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook containing a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first PivotTable in the worksheet
                PivotTable pivot = sheet.PivotTables[0];

                // Recalculate the PivotTable data to ensure it is up‑to‑date
                pivot.CalculateData();
            }
            else
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
            }

            // Save the workbook after recalculating the PivotTable
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
