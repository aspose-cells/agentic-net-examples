// Title: Refresh and recalculate every pivot table in an XLSX workbook using Aspose.Cells for .NET while preserving original content
// Common Searches: how to programmatically refresh pivot tables in an Excel file using Aspose.Cells C# | C# example to loop through worksheets and refresh all pivot tables with Aspose.Cells | update pivot table data source and recalculate after turning off Excel 2003 compatibility in .NET | Aspose.Cells .NET preserve original workbook while refreshing pivot tables | refresh pivot tables in XLSX workbook and save changes using Aspose.Cells for C#
// Tags: Aspose.Cells refresh pivot tables .NET | C# iterate worksheets to update pivot tables | PivotTable RefreshData CalculateData API | preserve original workbook content Aspose.Cells | disable Excel 2003 compatibility handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace AsposeCellsExample
{
    // The C# program loads an XLSX workbook with Aspose.Cells, iterates each worksheet, checks for pivot tables, calls RefreshData() and CalculateData() on each PivotTable to update its source and recalculate values, then saves the workbook, ensuring the original content remains intact.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Refresh all pivot tables in each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Ensure the worksheet contains pivot tables
                        if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
                            continue;

                        foreach (PivotTable pivotTable in sheet.PivotTables)
                        {
                            // Refresh the data source of the pivot table
                            pivotTable.RefreshData();

                            // Recalculate the pivot table after refreshing
                            pivotTable.CalculateData();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing pivot tables in sheet '{sheet.Name}': {ex.Message}");
                    }
                }

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
}
