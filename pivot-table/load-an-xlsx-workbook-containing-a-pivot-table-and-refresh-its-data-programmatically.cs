// Title: Refresh all pivot tables and recalculate formulas in an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh pivot tables in existing workbook | How to programmatically update pivot cache and recalculate data with Aspose.Cells .NET | Refresh pivot tables and recalculate formulas in Excel file using Aspose.Cells C# example | Iterate through worksheets to refresh all pivot tables Aspose.Cells .NET | Load XLSX, refresh pivot data and save with Aspose.Cells in C#
// Tags: Aspose.Cells pivot data refresh method | C# recalculate workbook formulas Aspose.Cells | Update XLSX pivot cache programmatically | Loop worksheets to process pivots Aspose.Cells | Export workbook after pivot update .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The C# program verifies the input XLSX file, loads it with Aspose.Cells, iterates over each worksheet and its PivotTables to call RefreshData and CalculateData, optionally recalculates all workbook formulas, and saves the modified workbook to a new file, handling any exceptions that occur.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Refresh all pivot tables in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    pivot.RefreshData();   // Refresh source data
                    pivot.CalculateData(); // Recalculate pivot values
                }
            }

            // Recalculate any formulas in the workbook (optional)
            workbook.CalculateFormula();

            // Save the workbook with refreshed pivot tables
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
