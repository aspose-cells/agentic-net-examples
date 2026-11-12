// Title: Programmatically refresh all pivot tables in an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh pivot tables in an existing Excel workbook with Aspose.Cells | Aspose.Cells .NET refresh pivot cache and recalculate formulas in XLSX | Update all pivot tables after source data changes using Aspose.Cells for .NET | Iterate worksheets and refresh each pivot table with Aspose.Cells C# example | Refresh pivot tables and recalculate workbook formulas programmatically in C#
// Tags: Aspose.Cells pivot table cache update | C# iterate worksheets Aspose.Cells | refresh pivot data Aspose.Cells .NET | calculate pivot data Aspose.Cells | recalculate workbook formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace AsposeCellsExample
{
    // The example verifies the input file, loads the XLSX workbook with Aspose.Cells, iterates through every worksheet and its PivotTable collection, calls RefreshData and CalculateData on each pivot to update the source data and cache, recalculates all workbook formulas, and saves the refreshed workbook to the specified output path.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook that contains pivot tables.
                Workbook workbook = new Workbook(inputPath);

                // Refresh all pivot tables in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        try
                        {
                            pivot.RefreshData();      // Refresh source data.
                            pivot.CalculateData();    // Recalculate the pivot cache.
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error refreshing pivot table '{pivot.Name}' on sheet '{sheet.Name}': {ex.Message}");
                        }
                    }
                }

                // Recalculate formulas in case the pivot refresh affects dependent calculations.
                workbook.CalculateFormula();

                // Save the workbook with the refreshed pivot tables.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
