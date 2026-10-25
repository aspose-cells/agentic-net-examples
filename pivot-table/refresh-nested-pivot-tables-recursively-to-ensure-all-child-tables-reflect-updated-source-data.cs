// Title: How to recursively refresh nested pivot tables in an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to refresh all pivot tables in an Excel file with Aspose.Cells | update dependent nested pivot tables after data change using Aspose.Cells .NET | recursive pivot table refresh example with Aspose.Cells workbook | Aspose.Cells refresh pivot tables twice to handle child tables in C# | load workbook, refresh nested pivot tables, and save using Aspose.Cells
// Tags: recursive pivot table refresh Aspose.Cells | refresh nested pivot tables .NET | update dependent pivot tables Excel C# | Aspose.Cells refresh data cache pivot | refresh all pivot tables workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel workbook with Aspose.Cells, iterates through each worksheet to refresh every pivot table, performs a second pass to ensure child (nested) pivot tables are updated, and then saves the modified workbook, handling file‑not‑found and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputFile);

            // Refresh all pivot tables, handling nested tables recursively
            RefreshAllPivotTables(workbook);

            // Save the updated workbook (lifecycle rule: save)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Recursively refreshes pivot tables in the entire workbook
    static void RefreshAllPivotTables(Workbook workbook)
    {
        // First pass: refresh every pivot table
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pt in sheet.PivotTables)
            {
                try
                {
                    // Refresh the data cache of the pivot table
                    pt.RefreshData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to refresh pivot table '{pt.Name}' on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Second pass: ensures that child (nested) pivot tables that depend on the results
        // of other pivot tables are also updated after the first pass
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pt in sheet.PivotTables)
            {
                try
                {
                    pt.RefreshData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to re‑refresh pivot table '{pt.Name}' on sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }
    }
}
