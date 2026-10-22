// Title: Refresh all pivot tables in an Excel workbook after modifying the data source using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, iterate through each worksheet and each PivotTable, call RefreshData() and CalculateData() on the pivot, then save the workbook to a new file. | Update a specific cell value in a worksheet, then programmatically refresh and recalculate every pivot table in the workbook using the Aspose.Cells C# API.
// Common Searches: C# Aspose.Cells how to refresh every pivot table after changing source data | Refresh all PivotTables in an Excel file using Aspose.Cells for .NET | Aspose.Cells RefreshData and CalculateData for multiple worksheets example | Programmatically recalculate pivot tables in a workbook with Aspose.Cells C# | Update cell value and refresh pivot tables in the same Aspose.Cells operation
// Tags: Aspose.Cells RefreshData pivot tables | Aspose.Cells CalculateData usage | C# refresh all pivots in workbook | update Excel data source before pivot refresh Aspose | iterate worksheets Aspose.Cells pivot refresh | save workbook after pivot refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefresh
{
    // The sample loads an existing Excel workbook, optionally modifies a cell value, then loops through every worksheet and each PivotTable to invoke RefreshData() and CalculateData(). After refreshing all pivots, the workbook is saved to a new file, with error handling for missing input and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // (Optional) Update the underlying data source here
                // Example: workbook.Worksheets[0].Cells["A2"].PutValue(123);

                // Refresh all pivot tables in every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (PivotTable pivot in sheet.PivotTables)
                    {
                        // Refresh the data source for the pivot table
                        pivot.RefreshData();

                        // Recalculate the pivot table values after refresh
                        pivot.CalculateData();
                    }
                }

                // Save the workbook with refreshed pivot tables
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
