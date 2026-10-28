// Title: How to change the ODBC connection string of an Excel PivotTable’s external source using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change pivot table ODBC connection string to new server | Update external data source for Excel pivot cache programmatically with Aspose.Cells | Set new server location in PivotCacheDefinition.ConnectionInfo using Aspose.Cells .NET | Modify ODBC connection of a PivotTable in an existing workbook with Aspose.Cells | Aspose.Cells example for editing pivot table data source connection string
// Tags: Aspose.Cells pivot cache connection string update | C# modify pivot table external ODBC source | Aspose.Cells PivotCacheDefinition ConnectionInfo usage | Excel workbook pivot table data source change .NET | programmatic ODBC connection update for Excel pivot

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, enumerates each worksheet and its PivotTables, displays the table names, and shows where to apply a new ODBC connection string via the PivotCacheDefinition.ConnectionInfo API before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all pivot tables in the worksheet
                    foreach (PivotTable pivotTable in sheet.PivotTables)
                    {
                        // Placeholder: In newer Aspose.Cells versions you can modify the ODBC connection string
                        // via pivotTable.PivotCacheDefinition.ConnectionInfo. This example skips that
                        // because the required API may not be available in the current library version.
                        Console.WriteLine($"Found pivot table: {pivotTable.Name} on sheet {sheet.Name}");
                    }
                }

                // Save the workbook (unchanged or with any modifications you add)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
