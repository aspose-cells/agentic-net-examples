// Title: Recalculate all PivotTables in an Excel workbook with Aspose.Cells for .NET before saving
// Common Searches: Aspose.Cells C# how to refresh pivot tables before saving workbook | calculate pivot table data programmatically using Aspose.Cells .NET | recalculate all pivot tables in an existing Excel file with Aspose.Cells | C# Aspose.Cells PivotTable.CalculateData example | update pivot cache in Excel workbook using Aspose.Cells before export | refresh pivot tables after modifying source data with Aspose.Cells
// Tags: Aspose.Cells PivotTable.CalculateData usage | refresh pivot tables .NET Excel | iterate worksheets pivot tables C# | calculate pivot data before workbook save | programmatic pivot cache refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotTableCollection

// The sample loads an existing Excel file with Aspose.Cells, iterates through each worksheet and its PivotTableCollection, calls PivotTable.CalculateData() on every pivot table to ensure the cached data is up‑to‑date, and then saves the workbook to a new file while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of pivot tables on the current worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Recalculate each pivot table's data
                foreach (PivotTable pivot in pivotTables)
                {
                    pivot.CalculateData();
                }
            }

            // Save the workbook after pivot data has been calculated
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
